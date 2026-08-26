import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { chmod, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const officeDir = path.dirname(fileURLToPath(import.meta.url));

async function executable(file, source) {
  await writeFile(file, `#!${process.execPath}\n${source}\n`, 'utf8');
  await chmod(file, 0o755);
}

function protocol(child) {
  const pending = new Map();
  let stdout = '';
  child.stdout.on('data', chunk => {
    stdout += chunk;
    for (;;) {
      const newline = stdout.indexOf('\n');
      if (newline < 0) break;
      const line = stdout.slice(0, newline).trim();
      stdout = stdout.slice(newline + 1);
      if (!line) continue;
      const message = JSON.parse(line);
      pending.get(message.id)?.(message);
      pending.delete(message.id);
    }
  });
  let nextId = 1;
  return (method, params = {}) => new Promise((resolve, reject) => {
    const id = nextId++;
    const timer = setTimeout(() => reject(new Error(`timeout waiting for ${method}`)), 10000);
    pending.set(id, message => {
      clearTimeout(timer);
      resolve(message);
    });
    child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  });
}

test('published Office MCP exposes bounded orthogonal capabilities', async () => {
  const temporary = await mkdtemp(path.join(os.tmpdir(), 'office-mcp-contract-'));
  const fake = `
const fs = require('node:fs');
const command = process.argv[2];
if (command === 'edit') {
  const input = process.argv[3];
  const operations = JSON.parse(fs.readFileSync(process.argv[4], 'utf8')).operations;
  const output = process.argv[5];
  fs.copyFileSync(input, output);
  process.stdout.write(JSON.stringify({
    input, output,
    appliedOperations: operations.map(operation => ({type: operation.type, applied: true, detail: 'ok'})),
  }));
  process.exit(0);
}
process.stderr.write('unexpected command: ' + command);
process.exit(2);
`;
  await Promise.all([
    executable(path.join(temporary, 'tiwater-docx'), fake),
    executable(path.join(temporary, 'tiwater-xlsx'), fake),
    executable(path.join(temporary, 'tiwater-pptx'), fake),
    executable(path.join(temporary, 'tiwater-convert'), fake),
  ]);
  const child = spawn(process.execPath, [path.join(officeDir, 'index.mjs')], {
    cwd: temporary,
    env: { ...process.env, PATH: `${temporary}${path.delimiter}${process.env.PATH}` },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  const request = protocol(child);
  try {
    const initialized = await request('initialize', {
      protocolVersion: '2025-06-18', capabilities: {},
      clientInfo: { name: 'office-contract', version: '1.0.0' },
    });
    assert.equal(initialized.result.serverInfo.version, '0.13.0');

    const listed = await request('tools/list');
    const names = listed.result.tools.map(tool => tool.name);
    for (const removed of ['docx_list_migration_choices', 'docx_query_migration_choices', 'docx_migrate_template', 'docx_verify_migration', 'xlsx_apply']) {
      assert(!names.includes(removed), `${removed} must not be public`);
    }
    for (const required of [
      'docx_inspect', 'docx_inspect_tables', 'docx_set_table_cell_text', 'docx_validate',
      'xlsx_convert_legacy', 'xlsx_set_cell_value', 'xlsx_set_page_setup', 'xlsx_validate',
      'pptx_inspect', 'pptx_apply_template', 'pptx_apply_format', 'pptx_validate',
      'office_render_pdf',
    ]) assert(names.includes(required), `missing ${required}`);

    assert(names.length >= 60 && names.length <= 100, `unexpected capability count: ${names.length}`);
    assert(!names.some(name => /scenario|migration|issue|workitem/i.test(name)));

    const tool = listed.result.tools.find(candidate => candidate.name === 'docx_set_table_cell_text');
    assert.deepEqual(tool.inputSchema.required.sort(), ['changes', 'input', 'output', 'receiptOutput']);
    assert.equal(tool.inputSchema.properties.changes.items.properties.type, undefined);
    assert.equal(tool.inputSchema.properties.changes.items.additionalProperties, false);

    const input = path.join(temporary, 'input.docx');
    const output = path.join(temporary, 'output.docx');
    const receiptOutput = path.join(temporary, 'receipt.json');
    await writeFile(input, 'current document', 'utf8');
    const edited = await request('tools/call', {
      name: 'docx_set_table_cell_text',
      arguments: { input, output, receiptOutput, changes: [{ tableIndex: 0, rowIndex: 1, cellIndex: 2, text: 'value' }] },
    });
    assert.equal(edited.result.structuredContent.summary.pass, true);
    assert.equal(edited.result.structuredContent.summary.operationCount, 1);
    const receipt = JSON.parse(await readFile(receiptOutput, 'utf8'));
    assert.equal(receipt.operationType, 'replaceTableCellText');
    assert.equal(receipt.appliedOperations[0].type, 'replaceTableCellText');
    assert.equal(receipt.input.path, input);
    assert.equal(receipt.output.path, output);

    const injected = await request('tools/call', {
      name: 'docx_set_table_cell_text',
      arguments: { input, output: path.join(temporary, 'bad.docx'), receiptOutput: path.join(temporary, 'bad.json'), changes: [{ tableIndex: 0, rowIndex: 0, cellIndex: 0, text: 'x', type: 'deleteBodyRange' }] },
    });
    assert.equal(injected.result.isError, true);
  } finally {
    if (child.exitCode === null) {
      child.kill('SIGTERM');
      await new Promise(resolve => child.once('exit', resolve));
    }
    await rm(temporary, { recursive: true, force: true });
  }
});
