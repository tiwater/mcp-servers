import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { chmod, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const pdfDir = path.dirname(fileURLToPath(import.meta.url));

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
    pending.set(id, message => { clearTimeout(timer); resolve(message); });
    child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  });
}

test('PDF MCP fixes OCR provider and model without credential arguments', async () => {
  const temporary = await mkdtemp(path.join(os.tmpdir(), 'pdf-mcp-contract-'));
  const runtime = path.join(temporary, 'tiwater-pdf');
  await writeFile(runtime, `#!${process.execPath}
const args = process.argv.slice(2);
if (args[0] !== 'ocr' || !args.includes('--provider') || !args.includes('llm') || !args.includes('--llm-model') || !args.includes('qwen3.8-max')) process.exit(2);
process.stdout.write(JSON.stringify({model:'qwen3.8-max',page_count:2,pages:[{page:1},{page:2}]}));
`, 'utf8');
  await chmod(runtime, 0o755);
  const child = spawn(process.execPath, [path.join(pdfDir, 'index.mjs')], {
    cwd: temporary,
    env: { ...process.env, PATH: `${temporary}${path.delimiter}${process.env.PATH}` },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  const request = protocol(child);
  try {
    await request('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'pdf-contract', version: '1.0.0' } });
    const listed = await request('tools/list');
    const byName = new Map(listed.result.tools.map(tool => [tool.name, tool]));
    assert(byName.has('pdf_ocr'));
    for (const name of ['pdf_extract_tables', 'pdf_find_table', 'pdf_ocr']) {
      const properties = byName.get(name).inputSchema.properties;
      for (const forbidden of ['apiKey', 'baseUrl', 'llmModel', 'llmFallback', 'provider']) assert.equal(properties[forbidden], undefined);
    }
    const output = path.join(temporary, 'ocr.json');
    const called = await request('tools/call', { name: 'pdf_ocr', arguments: { input: '/current.pdf', output, pages: [1, 2] } });
    assert.equal(called.result.structuredContent.summary.model, 'qwen3.8-max');
    assert.equal(called.result.structuredContent.summary.pageCount, 2);
    assert.equal(JSON.parse(await readFile(output, 'utf8')).model, 'qwen3.8-max');
    const rejected = await request('tools/call', { name: 'pdf_ocr', arguments: { input: '/current.pdf', output: path.join(temporary, 'bad.json'), apiKey: 'forbidden' } });
    assert.equal(rejected.result?.isError ?? rejected.error?.code !== undefined, true, JSON.stringify(rejected));
  } finally {
    if (child.exitCode === null) {
      child.kill('SIGTERM');
      await new Promise(resolve => child.once('exit', resolve));
    }
    await rm(temporary, { recursive: true, force: true });
  }
});
