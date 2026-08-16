import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { chmod, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const officeDir = path.dirname(fileURLToPath(import.meta.url));

test('published Office MCP exposes one batch migration surface and forwards typed choices', async () => {
  const temporary = await mkdtemp(path.join(os.tmpdir(), 'office-mcp-test-'));
  const fakeRuntime = path.join(temporary, 'tiwater-docx');
  const fakeDotnet = path.join(temporary, 'dotnet');
  await writeFile(fakeDotnet, '#!/bin/sh\nexit 0\n', 'utf8');
  await chmod(fakeDotnet, 0o755);
  await writeFile(fakeRuntime, `#!${process.execPath}
const fs = require('node:fs');
const command = process.argv[2];
if (command === 'inspect') {
  process.stdout.write(JSON.stringify({schema:'inspection/v1',file:process.argv[3],tables:[{rows:2}],dotnetRoot:process.env.DOTNET_ROOT,dotnetRootArm64:process.env.DOTNET_ROOT_ARM64}));
  process.exit(0);
}
if (command === 'list-template-migration-choices') {
  const choice = id => ({id,kind:'paragraph',scope:'body',text:id,count:1,requiredCardinality:'one',context:null,allowedActions:['place-content']});
  const sources = process.argv[3] === '/invalid-output.docx' ? [{id:'source-1'}] : [choice('source-1')];
  process.stdout.write(JSON.stringify({schema:'catalog/v1',pass:true,sourceSha256:'a',baselineSha256:'b',sources,targets:[choice('target-1')]}));
  process.exit(0);
}
if (command === 'migrate-template' || command === 'verify-template-migration') {
  const payload = JSON.parse(fs.readFileSync(process.argv[5], 'utf8'));
  process.stdout.write(JSON.stringify({schema:'receipt/v1',toolVersion:'0.12.2',status:'pass',pass:true,reviewRequired:false,outputVerified:true,command,payload,output:process.argv[6],plan:process.argv[6]+'.migration-plan.json',failures:[]}));
  process.exit(0);
}
process.stderr.write('unexpected command: ' + command);
process.exit(2);
`, 'utf8');
  await chmod(fakeRuntime, 0o755);

  const child = spawn(process.execPath, [path.join(officeDir, 'index.mjs')], {
    env: { ...process.env, PATH: temporary, DOTNET_ROOT: '', DOTNET_ROOT_ARM64: '', DOTNET_ROOT_X64: '' },
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
    const initialized = await request('initialize', {
      protocolVersion: '2025-06-18',
      capabilities: {},
      clientInfo: { name: 'office-mcp-contract-test', version: '1.0.0' },
    });
    assert.equal(initialized.result.serverInfo.version, '0.3.1');
    const listed = await request('tools/list');
    const names = listed.result.tools.map(tool => tool.name);
    assert(names.includes('docx_list_migration_choices'));
    assert(names.includes('docx_migrate_template'));
    assert(names.includes('docx_verify_migration'));
    assert.deepEqual(
      names.filter(name => [
        'docx_edit',
        'docx_fill_template',
        'docx_strip_direct_formatting',
        'docx_replace_style_ids',
        'xlsx_edit',
        'xlsx_fill_template',
        'pptx_apply_format_edits',
        'pptx_fill_template',
      ].includes(name)),
      [],
    );
    assert(!names.some(name => name.includes('record') || name.includes('revise') || name.includes('target_search')));
    assert(!names.includes('docx_inspect_tables'));
    assert(!names.includes('pptx_inspect_detail'));
    const migrateTool = listed.result.tools.find(tool => tool.name === 'docx_migrate_template');
    const listTool = listed.result.tools.find(tool => tool.name === 'docx_list_migration_choices');
    assert.equal(listTool.outputSchema.properties.catalog.properties.sources.items.properties.allowedActions.type, 'array');
    assert.equal(migrateTool.outputSchema.properties.receipt.properties.outputVerified.type, 'boolean');
    assert.deepEqual(migrateTool.inputSchema.properties.choices.items.properties.action.enum, [
      'place-content',
      'keep-template-content',
      'keep-template-label',
      'select-template-option',
      'exclude-source',
      'review-source',
    ]);
    assert.equal(migrateTool.inputSchema.additionalProperties, false);

    const invalid = await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', invented: true },
    });
    assert.equal(invalid.result.isError, true);

    const invalidOutput = await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/invalid-output.docx', baseline: '/baseline.docx' },
    });
    assert.equal(invalidOutput.result.isError, true);

    const observationPath = path.join(temporary, 'observations', 'current.docx.json');
    const observed = await request('tools/call', {
      name: 'docx_inspect',
      arguments: { input: '/current.docx', output: observationPath },
    });
    assert.equal(observed.result.structuredContent.artifact.path, observationPath);
    assert.match(observed.result.structuredContent.artifact.sha256, /^[0-9a-f]{64}$/);
    assert.deepEqual(JSON.parse(await readFile(observationPath, 'utf8')), {
      schema: 'inspection/v1',
      file: '/current.docx',
      tables: [{ rows: 2 }],
      dotnetRoot: temporary,
      ...(process.arch === 'arm64' ? { dotnetRootArm64: temporary } : {}),
    });
    const overwrite = await request('tools/call', {
      name: 'docx_inspect',
      arguments: { input: '/current.docx', output: observationPath },
    });
    assert.equal(overwrite.result.isError, true);

    const listedChoices = await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx' },
    });
    assert.equal(listedChoices.result.structuredContent.runtime.cwd, process.cwd());
    assert.equal(listedChoices.result.structuredContent.catalog.sources[0].id, 'source-1');

    const choices = [{ sourceChoiceId: 'source-1', action: 'place-content', targetChoiceId: 'target-1' }];
    const migrated = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', output: '/output.docx', choices },
    });
    const payload = JSON.parse(migrated.result.content[0].text);
    assert.deepEqual(migrated.result.structuredContent, payload);
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
