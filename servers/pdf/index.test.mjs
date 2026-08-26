import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { chmod, mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

const serverPath = new URL('./index.mjs', import.meta.url).pathname;

test('PDF MCP fixes OCR to qwen3.8-max and does not expose arbitrary LLM table fallback', async () => {
  const binDir = await mkdtemp(path.join(os.tmpdir(), 'tiwater-pdf-test-'));
  const fakeTool = path.join(binDir, 'tiwater-pdf');
  await writeFile(fakeTool, '#!/usr/bin/env node\nprocess.stdout.write(JSON.stringify({ args: process.argv.slice(2) }));\n', 'utf8');
  await chmod(fakeTool, 0o755);
  const child = spawn(process.execPath, [serverPath], {
    env: { ...process.env, PATH: `${binDir}:${process.env.PATH}` },
    stdio: ['pipe', 'pipe', 'inherit'],
  });
  let nextId = 1;
  let buffer = '';
  const pending = new Map();
  child.stdout.on('data', chunk => {
    buffer += chunk.toString();
    const lines = buffer.split('\n');
    buffer = lines.pop();
    for (const line of lines) {
      const message = JSON.parse(line);
      const resolve = pending.get(message.id);
      if (resolve) {
        pending.delete(message.id);
        resolve(message);
      }
    }
  });
  const request = (method, params = {}) => new Promise(resolve => {
    const id = nextId++;
    pending.set(id, resolve);
    child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  });

  try {
    const listed = await request('tools/list');
    const byName = new Map(listed.result.tools.map(tool => [tool.name, tool]));
    assert.equal(byName.has('pdf_ocr'), true);
    for (const name of ['pdf_extract_tables', 'pdf_find_table']) {
      const properties = byName.get(name).inputSchema.properties;
      assert.equal('apiKey' in properties, false);
      assert.equal('llmModel' in properties, false);
      assert.equal('llmFallback' in properties, false);
    }

    const called = await request('tools/call', {
      name: 'pdf_ocr',
      arguments: { input: '/tmp/input.pdf', pages: [2, 1] },
    });
    const payload = JSON.parse(called.result.content[0].text);
    assert.deepEqual(payload.report.args, [
      'ocr', '/tmp/input.pdf', '--provider', 'llm', '--llm-model', 'qwen3.8-max', '--pages', '2,1', '--json',
    ]);
    assert.equal(JSON.stringify(payload).includes('api-key'), false);
  } finally {
    child.stdin.end();
    await new Promise(resolve => child.once('exit', resolve));
    await rm(binDir, { recursive: true, force: true });
  }
});
