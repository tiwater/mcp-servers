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
const choice = (id, options={}) => ({id,kind:options.kind??'paragraph',scope:options.scope??'body',text:options.text??id,count:1,requiredCardinality:options.requiredCardinality??'one',context:options.context??null,allowedActions:options.allowedActions??['place-content','keep-template-content','keep-template-label']});
const catalogFor = sourcePath => {
  let sources = [choice('source-1')];
  let targets = [choice('target-1')];
  if (sourcePath === '/invalid-output.docx') sources = [{id:'source-1'}];
  if (sourcePath === '/empty.docx') { sources = []; targets = []; }
  if (sourcePath === '/multi.docx') {
    sources = [
      choice('source-alpha', {text:'Alpha heading'}),
      choice('source-beta', {text:'Beta value',kind:'table-cell'}),
    ];
    targets = [
      choice('target-header', {text:'Current heading',scope:'header'}),
      choice('target-revision', {text:'',kind:'table-cell',context:{sameRowTexts:['Revision history','01']},allowedActions:['place-content','keep-template-content','keep-template-label','template-cleanup']}),
      choice('target-body', {text:'Beta destination',kind:'table-cell'}),
    ];
  }
  if (sourcePath === '/many.docx') {
    targets = Array.from({length:25}, (_, index) => choice('target-' + String(index + 1).padStart(2, '0')));
  }
  return {schema:'catalog/v1',pass:true,sourceSha256:'a',baselineSha256:'b',sources,targets};
};
if (command === 'inspect') {
  process.stdout.write(JSON.stringify({schema:'inspection/v1',file:process.argv[3],tables:[{rows:2}],dotnetRoot:process.env.DOTNET_ROOT,dotnetRootArm64:process.env.DOTNET_ROOT_ARM64}));
  process.exit(0);
}
if (command === 'list-template-migration-choices') {
  process.stdout.write(JSON.stringify(catalogFor(process.argv[3])));
  process.exit(0);
}
if (command === 'find-template-migration-targets') {
  const catalog = catalogFor(process.argv[3]);
  const sourceChoiceId = process.argv[5] === '-' ? null : process.argv[5];
  const branch = process.argv[6];
  const query = process.argv[7] === '-' ? null : process.argv[7].toLowerCase();
  const offset = Number(process.argv[8]);
  const limit = Number(process.argv[9]);
  const source = sourceChoiceId ? catalog.sources.find(item => item.id === sourceChoiceId) : null;
  let targets = catalog.targets.filter(item => {
    if (branch === 'baseline-clear') return item.allowedActions.includes('template-cleanup');
    if (branch === 'choice-selection') return item.allowedActions.includes('select-template-option');
    if (!source || item.kind !== source.kind) return false;
    if (branch === 'copy-text' || branch === 'copy-media') return item.allowedActions.includes('place-content');
    if (branch === 'retain-target') return item.allowedActions.includes('keep-template-content');
    if (branch === 'retain-target-label') return item.allowedActions.includes('keep-template-label');
    return false;
  });
  if (query) targets = targets.filter(item => JSON.stringify({text:item.text,context:item.context}).toLowerCase().includes(query));
  if (process.argv[3] === '/drift.docx') targets = [choice('target-not-in-current-catalog')];
  targets.sort((left,right)=>left.id.localeCompare(right.id));
  process.stdout.write(JSON.stringify({schema:'target-page/v1',pass:true,sourceChoiceId,branch,offset,limit,total:targets.length,targets:targets.slice(offset,offset+limit)}));
  process.exit(0);
}
if (command === 'migrate-template' || command === 'verify-template-migration') {
  const payload = JSON.parse(fs.readFileSync(process.argv[5], 'utf8'));
  const failed = process.argv[6] === '/failed.docx';
  process.stdout.write(JSON.stringify({schema:'receipt/v1',toolVersion:'0.12.2',status:failed?'failed':'pass',pass:!failed,reviewRequired:false,outputVerified:!failed,command,payload,output:failed?null:process.argv[6],plan:failed?null:process.argv[6]+'.migration-plan.json',failures:failed?[{reason:'known-failure'}]:[]}));
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
    assert.equal(initialized.result.serverInfo.version, '0.6.0');
    const listed = await request('tools/list');
    const names = listed.result.tools.map(tool => tool.name);
    assert(names.includes('docx_list_migration_choices'));
    assert(names.includes('docx_query_migration_choices'));
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
    assert.equal(listTool.inputSchema.required.includes('output'), true);
    assert.equal(listTool.outputSchema.properties.artifact.properties.sha256.type, 'string');
    assert.equal(listTool.outputSchema.properties.summary.properties.sourceCount.type, 'integer');
    assert.equal(migrateTool.inputSchema.required.includes('receiptOutput'), true);
    assert.equal(migrateTool.outputSchema.properties.artifact.properties.sha256.type, 'string');
    assert.equal(migrateTool.outputSchema.properties.summary.properties.outputVerified.type, 'boolean');
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
      arguments: { source: '/invalid-output.docx', baseline: '/baseline.docx', output: path.join(temporary, 'invalid.catalog.json') },
    });
    assert.equal(invalidOutput.result.isError, true);
    await assert.rejects(readFile(path.join(temporary, 'invalid.catalog.json'), 'utf8'));

    const observationPath = path.join(temporary, 'observations', 'current.docx.json');
    const observed = await request('tools/call', {
      name: 'docx_inspect',
      arguments: { input: '/current.docx', output: observationPath },
    });
    assert.equal(observed.result.structuredContent.artifact.path, observationPath);
    assert.match(observed.result.structuredContent.artifact.sha256, /^[0-9a-f]{64}$/);
    const observation = JSON.parse(await readFile(observationPath, 'utf8'));
    assert.deepEqual(observation, {
      schema: 'inspection/v1',
      file: '/current.docx',
      tables: [{ rows: 2 }],
      dotnetRoot: temporary,
      dotnetRootArm64: process.arch === 'arm64' ? temporary : '',
    });
    const overwrite = await request('tools/call', {
      name: 'docx_inspect',
      arguments: { input: '/current.docx', output: observationPath },
    });
    assert.equal(overwrite.result.isError, true);

    const catalogPath = path.join(temporary, 'catalogs', 'record.json');
    const listedChoices = await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', output: catalogPath },
    });
    assert.equal(listedChoices.result.structuredContent.runtime.cwd, process.cwd());
    assert.equal(listedChoices.result.structuredContent.artifact.path, catalogPath);
    assert.equal(listedChoices.result.structuredContent.summary.sourceCount, 1);
    assert.equal(listedChoices.result.structuredContent.summary.targetCount, 1);
    assert.equal('catalog' in listedChoices.result.structuredContent, false);
    assert.equal(JSON.parse(await readFile(catalogPath, 'utf8')).sources[0].id, 'source-1');

    const sourcePage = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'sources', offset: 0, limit: 10 },
    });
    assert.equal(sourcePage.result.structuredContent.view, 'sources');
    assert.deepEqual(sourcePage.result.structuredContent.items.map(item => item.id), ['source-1']);
    assert.deepEqual(sourcePage.result.structuredContent.page, {
      offset: 0, returned: 1, total: 1, hasMore: false,
    });
    assert.equal(sourcePage.result.structuredContent.sourceSha256, 'a');
    assert.equal(sourcePage.result.structuredContent.baselineSha256, 'b');

    const targetPage = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: {
        source: '/current.docx',
        baseline: '/baseline.docx',
        view: 'targets',
        sourceChoiceId: 'source-1',
        action: 'place-content',
        text: 'TARGET',
        offset: 0,
        limit: 10,
      },
    });
    assert.equal(targetPage.result.structuredContent.view, 'targets');
    assert.equal(targetPage.result.structuredContent.action, 'place-content');
    assert.equal(targetPage.result.structuredContent.source.id, 'source-1');
    assert.deepEqual(targetPage.result.structuredContent.items.map(item => item.id), ['target-1']);
    assert.deepEqual(targetPage.result.structuredContent.page, {
      offset: 0, returned: 1, total: 1, hasMore: false,
    });

    const unknownSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'targets', sourceChoiceId: 'missing-source', action: 'place-content' },
    });
    assert.equal(unknownSource.result.isError, true);

    const oversizedPage = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'sources', limit: 11 },
    });
    assert.equal(oversizedPage.result.isError, true);

    const multiCatalogPath = path.join(temporary, 'catalogs', 'multi.json');
    await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', output: multiCatalogPath },
    });
    const firstSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'sources', limit: 1 },
    });
    assert.deepEqual(firstSource.result.structuredContent.items.map(item => item.id), ['source-alpha']);
    assert.equal(firstSource.result.structuredContent.page.hasMore, true);
    const secondSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'sources', offset: 1, limit: 1 },
    });
    assert.deepEqual(secondSource.result.structuredContent.items.map(item => item.id), ['source-beta']);
    assert.equal(secondSource.result.structuredContent.page.hasMore, false);

    const contextMatch = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: {
        source: '/multi.docx',
        baseline: '/baseline.docx',
        view: 'targets',
        sourceChoiceId: 'source-beta',
        action: 'place-content',
        text: 'REVISION',
      },
    });
    assert.deepEqual(contextMatch.result.structuredContent.items.map(item => item.id), ['target-revision']);
    const noMatches = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'targets', sourceChoiceId: 'source-beta', action: 'place-content', text: 'not present' },
    });
    assert.deepEqual(noMatches.result.structuredContent.items, []);
    assert.equal(noMatches.result.structuredContent.page.total, 0);
    const opaqueIdIsNotVisibleText = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'targets', sourceChoiceId: 'source-beta', action: 'place-content', text: 'target-header' },
    });
    assert.deepEqual(opaqueIdIsNotVisibleText.result.structuredContent.items, []);

    const cleanupTargets = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'cleanup', text: 'revision' },
    });
    assert.equal(cleanupTargets.result.structuredContent.source, null);
    assert.equal(cleanupTargets.result.structuredContent.action, null);
    assert.deepEqual(cleanupTargets.result.structuredContent.items.map(item => item.id), ['target-revision']);

    const manyCatalogPath = path.join(temporary, 'catalogs', 'many.json');
    await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/many.docx', baseline: '/baseline.docx', output: manyCatalogPath },
    });
    const boundedTargets = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/many.docx', baseline: '/baseline.docx', view: 'targets', sourceChoiceId: 'source-1', action: 'place-content', limit: 10 },
    });
    assert.equal(boundedTargets.result.structuredContent.items.length, 10);
    assert.equal(boundedTargets.result.structuredContent.page.total, 25);
    assert.equal(boundedTargets.result.structuredContent.page.hasMore, true);

    const invalidCatalogSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/invalid-output.docx', baseline: '/baseline.docx', view: 'sources' },
    });
    assert.equal(invalidCatalogSource.result.isError, true);

    const actionNotAllowed = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'targets', sourceChoiceId: 'source-1', action: 'select-template-option' },
    });
    assert.equal(actionNotAllowed.result.isError, true);

    const targetDrift = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/drift.docx', baseline: '/baseline.docx', view: 'targets', sourceChoiceId: 'source-1', action: 'place-content' },
    });
    assert.equal(targetDrift.result.isError, true);

    const inventedQueryField = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'sources', recommendation: true },
    });
    assert.equal(inventedQueryField.result.isError, true);

    const relocatedCatalogPath = path.join(temporary, 'relocated', 'record.json');
    const relocated = await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', output: relocatedCatalogPath },
    });
    assert.deepEqual(relocated.result.structuredContent.summary, listedChoices.result.structuredContent.summary);
    assert.deepEqual(
      JSON.parse(await readFile(relocatedCatalogPath, 'utf8')),
      JSON.parse(await readFile(catalogPath, 'utf8')),
    );

    const emptyCatalogPath = path.join(temporary, 'catalogs', 'empty.json');
    const empty = await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/empty.docx', baseline: '/baseline.docx', output: emptyCatalogPath },
    });
    assert.equal(empty.result.structuredContent.summary.sourceCount, 0);
    assert.equal(empty.result.structuredContent.summary.targetCount, 0);
    assert.deepEqual(JSON.parse(await readFile(emptyCatalogPath, 'utf8')).sources, []);

    const overwriteCatalog = await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', output: catalogPath },
    });
    assert.equal(overwriteCatalog.result.isError, true);

    const choices = [{ sourceChoiceId: 'source-1', action: 'place-content', targetChoiceId: 'target-1' }];
    const migrationReceiptPath = path.join(temporary, 'receipts', 'migration.json');
    const migrated = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/output.docx',
        receiptOutput: migrationReceiptPath, choices,
      },
    });
    const payload = JSON.parse(migrated.result.content[0].text);
    assert.deepEqual(migrated.result.structuredContent, payload);
    assert.equal(payload.summary.pass, true);
    assert.equal(payload.summary.failureCount, 0);
    assert.equal('receipt' in payload, false);
    const migrationReceipt = JSON.parse(await readFile(migrationReceiptPath, 'utf8'));
    assert.equal(migrationReceipt.command, 'migrate-template');
    assert.deepEqual(migrationReceipt.payload, {
      schema: 'tiwater.docx.template-migration-business-choices/v1',
      choices,
    });

    const verificationReceiptPath = path.join(temporary, 'receipts', 'verification.json');
    const verified = await request('tools/call', {
      name: 'docx_verify_migration',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/output.docx',
        receiptOutput: verificationReceiptPath, choices,
      },
    });
    assert.equal(verified.result.structuredContent.summary.pass, true);
    assert.equal(verified.result.structuredContent.summary.outputVerified, true);
    assert.equal('receipt' in verified.result.structuredContent, false);
    assert.equal(JSON.parse(await readFile(verificationReceiptPath, 'utf8')).command, 'verify-template-migration');

    const failedReceiptPath = path.join(temporary, 'receipts', 'failed.json');
    const failed = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/failed.docx',
        receiptOutput: failedReceiptPath, choices,
      },
    });
    assert.equal(failed.result.isError, undefined);
    assert.equal(failed.result.structuredContent.summary.status, 'failed');
    assert.equal(failed.result.structuredContent.summary.failureCount, 1);
    assert.deepEqual(JSON.parse(await readFile(failedReceiptPath, 'utf8')).failures, [{ reason: 'known-failure' }]);

    const receiptOverwrite = await request('tools/call', {
      name: 'docx_verify_migration',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/output.docx',
        receiptOutput: migrationReceiptPath, choices,
      },
    });
    assert.equal(receiptOverwrite.result.isError, true);
  } finally {
    if (child.exitCode === null && child.signalCode === null) {
      child.kill('SIGTERM');
      await new Promise(resolve => child.once('exit', resolve));
    }
    await rm(temporary, { recursive: true, force: true });
  }
});
