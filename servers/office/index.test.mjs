import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { chmod, mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

const serverPath = new URL('./index.mjs', import.meta.url).pathname;

test('office MCP exposes fixed-action edit tools and fixes the provider operation type', async () => {
  const binDir = await mkdtemp(path.join(os.tmpdir(), 'tiwater-office-test-'));
  const fakeTool = path.join(binDir, 'fake-office-tool');
  await writeFile(fakeTool, `#!/usr/bin/env node
const fs = require('node:fs');
const args = process.argv.slice(2);
const operations = args[0] === 'edit' ? JSON.parse(fs.readFileSync(args[2], 'utf8')).operations : [];
process.stdout.write(JSON.stringify({ args, operations }));
`, 'utf8');
  await chmod(fakeTool, 0o755);
  for (const name of ['tiwater-docx', 'tiwater-xlsx', 'tiwater-pptx']) {
    await writeFile(path.join(binDir, name), `#!/bin/sh\nexec "${fakeTool}" "$@"\n`, 'utf8');
    await chmod(path.join(binDir, name), 0o755);
  }

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
      const entry = pending.get(message.id);
      if (entry) {
        pending.delete(message.id);
        entry(message);
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
    const names = listed.result.tools.map(tool => tool.name);
    assert.equal(names.includes('docx_edit'), false);
    assert.equal(names.includes('xlsx_edit'), false);
    assert.equal(names.includes('pptx_apply_format_edits'), false);
    assert.equal(names.includes('pptx_inspect_detail'), false);
    assert.equal(names.includes('docx_set_table_cell_text'), true);
    assert.equal(names.includes('xlsx_set_cell_value'), true);
    assert.equal(names.includes('pptx_set_text_format'), true);
    assert.equal(names.includes('pptx_apply_template'), true);
    assert.equal(names.includes('docx_validate'), true);
    assert.equal(names.includes('docx_validate_font_policy'), true);
    assert.equal(names.includes('pptx_validate'), true);

    const called = await request('tools/call', {
      name: 'docx_set_table_cell_text',
      arguments: {
        input: '/tmp/input.docx',
        output: '/tmp/output.docx',
        changes: [{ tableIndex: 0, rowIndex: 1, cellIndex: 2, text: 'value', type: 'deleteComments' }],
      },
    });
    const payload = JSON.parse(called.result.content[0].text);
    assert.deepEqual(payload.result.operations, [{ tableIndex: 0, rowIndex: 1, cellIndex: 2, text: 'value', type: 'replaceTableCellText' }]);

    const fieldCall = await request('tools/call', {
      name: 'docx_mark_fields_dirty',
      arguments: { input: '/tmp/input.docx', output: '/tmp/output.docx' },
    });
    const fieldPayload = JSON.parse(fieldCall.result.content[0].text);
    assert.deepEqual(fieldPayload.result.operations, [{ type: 'markFieldsDirty' }]);

    const templateCall = await request('tools/call', {
      name: 'pptx_apply_template',
      arguments: {
        input: '/tmp/input.pptx',
        template: '/tmp/template.pptx',
        output: '/tmp/output.pptx',
        targetMasterPath: 'ppt/slideMasters/slideMaster1.xml',
        slides: [{ slideNumber: 1, targetLayoutPath: 'ppt/slideLayouts/slideLayout1.xml' }],
      },
    });
    const templatePayload = JSON.parse(templateCall.result.content[0].text);
    assert.equal(templatePayload.result.args[0], 'apply-template');
    assert.equal(templatePayload.result.operations.length, 0);
  } finally {
    child.stdin.end();
    await new Promise(resolve => child.once('exit', resolve));
    await rm(binDir, { recursive: true, force: true });
  }
});
