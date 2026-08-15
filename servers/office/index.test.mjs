import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { chmod, mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const officeDir = path.dirname(fileURLToPath(import.meta.url));

test('published Office MCP exposes one batch migration surface and forwards typed choices', async () => {
  const temporary = await mkdtemp(path.join(os.tmpdir(), 'office-mcp-test-'));
  const fakeRuntime = path.join(temporary, 'tiwater-docx');
  await writeFile(fakeRuntime, `#!/usr/bin/env node
const fs = require('node:fs');
const command = process.argv[2];
if (command === 'list-template-migration-choices') {
  process.stdout.write(JSON.stringify({schema:'catalog/v1',sources:[{id:'source-1'}],targets:[{id:'target-1'}]}));
  process.exit(0);
}
if (command === 'migrate-template' || command === 'verify-template-migration') {
  const payload = JSON.parse(fs.readFileSync(process.argv[5], 'utf8'));
  process.stdout.write(JSON.stringify({schema:'receipt/v1',pass:true,command,payload,output:process.argv[6]}));
  process.exit(0);
}
process.stderr.write('unexpected command: ' + command);
process.exit(2);
`, 'utf8');
  await chmod(fakeRuntime, 0o755);

  const child = spawn(process.execPath, [path.join(officeDir, 'index.mjs')], {
    env: { ...process.env, PATH: `${temporary}${path.delimiter}${process.env.PATH ?? ''}` },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  const pending = new Map();
  let stdout = '';
  child.stdout.on('data', chunk => {
    stdout += chunk.toString();
    while (stdout.includes('\n')) {
      const newline = stdout.indexOf('\n');
      const line = stdout.slice(0, newline).trim();
      stdout = stdout.slice(newline + 1);
      if (!line) continue;
      const message = JSON.parse(line);
      pending.get(message.id)?.(message);
      pending.delete(message.id);
    }
  });
  let nextId = 1;
  const request = (method, params = {}) => new Promise((resolve, reject) => {
    const id = nextId++;
    const timer = setTimeout(() => reject(new Error(`timeout waiting for ${method}`)), 5000);
    pending.set(id, message => {
      clearTimeout(timer);
      resolve(message);
    });
    child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  });

  try {
    await request('initialize', { protocolVersion: '2025-06-18' });
    const listed = await request('tools/list');
    const names = listed.result.tools.map(tool => tool.name);
    assert(names.includes('docx_list_migration_choices'));
    assert(names.includes('docx_migrate_template'));
    assert(names.includes('docx_verify_migration'));
    assert.deepEqual(
      names.filter(name => ['docx_edit', 'docx_fill_template', 'xlsx_edit', 'xlsx_fill_template', 'pptx_apply_format_edits', 'pptx_fill_template'].includes(name)),
      [],
    );
    assert(!names.some(name => name.includes('record') || name.includes('revise') || name.includes('target_search')));
    const migrateTool = listed.result.tools.find(tool => tool.name === 'docx_migrate_template');
    assert.deepEqual(migrateTool.inputSchema.properties.choices.items.properties.action.enum, [
      'place-content',
      'keep-template-content',
      'keep-template-label',
      'select-template-option',
      'exclude-source',
      'review-source',
    ]);
    assert.equal(migrateTool.inputSchema.additionalProperties, false);

    const choices = [{ sourceChoiceId: 'source-1', action: 'place-content', targetChoiceId: 'target-1' }];
    const migrated = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', output: '/output.docx', choices },
    });
    const payload = JSON.parse(migrated.result.content[0].text);
    assert.equal(payload.receipt.pass, true);
    assert.equal(payload.receipt.command, 'migrate-template');
    assert.deepEqual(payload.receipt.payload, {
      schema: 'tiwater.docx.template-migration-business-choices/v1',
      choices,
    });
  } finally {
    child.kill('SIGTERM');
    await new Promise(resolve => child.once('exit', resolve));
    await rm(temporary, { recursive: true, force: true });
  }
});
