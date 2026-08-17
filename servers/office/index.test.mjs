import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { chmod, mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
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
const crypto = require('node:crypto');
const command = process.argv[2];
const choice = (id, options={}) => ({id,kind:options.kind??'paragraph',scope:options.scope??'body',text:options.text??id,count:1,requiredCardinality:options.requiredCardinality??'one',context:options.context??null,allowedActions:options.allowedActions??['place-content','keep-template-content','keep-template-label']});
const catalogFor = sourcePath => {
  let sources = [choice('source-1')];
  let targets = [choice('target-1')];
  if (sourcePath === '/invalid-output.docx') sources = [{id:'source-1'}];
  if (sourcePath === '/empty.docx') { sources = []; targets = []; }
  if (sourcePath === '/all.docx') {
    sources = [choice('source-all', {requiredCardinality:'all',allowedActions:['exclude-source','review-source']})];
  }
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
  if (sourcePath === '/context.docx' || sourcePath === '/context-mutated.docx') {
    sources = [choice('source-context', {text:'R-204',kind:'table-cell',context:{sameRowTexts:['2027-04-03','Changed scope'],columnHeaderText:'Release identifier',tableHeaderTexts:['Release identifier','Effective on','Change narrative']}})];
    targets = [
      choice('target-decoy', {text:'R-204',kind:'table-cell',context:{sameRowTexts:['Owner'],columnHeaderText:'Approver',tableHeaderTexts:['Approver','Signed on']}}),
      choice('target-context', {text:'template-value',kind:'table-cell',context:sourcePath === '/context.docx'?{sameRowTexts:['template-date','template-summary'],columnHeaderText:'Release identifier',tableHeaderTexts:['Release identifier','Effective on','Change narrative']}:{sameRowTexts:['Owner'],columnHeaderText:'Approver',tableHeaderTexts:['Approver','Signed on']}}),
    ];
  }
  if (sourcePath === '/many.docx') {
    targets = Array.from({length:25}, (_, index) => choice('target-' + String(index + 1).padStart(2, '0')));
  }
  return {schema:'catalog/v1',pass:true,sourceSha256:sourcePath==='/changed.docx'?'c':'a',baselineSha256:'b',sources,targets};
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
  if (process.argv[6] === '/silent.docx') {
    process.stderr.write('typed-migration-refusal');
    process.exit(1);
  }
  const failed = process.argv[6] === '/failed.docx';
  process.stdout.write(JSON.stringify({schema:'receipt/v1',toolVersion:'0.12.2',status:failed?'failed':'pass',pass:!failed,reviewRequired:false,outputVerified:!failed,command,payload,output:failed?null:process.argv[6],plan:failed?null:process.argv[6]+'.migration-plan.json',failures:failed?[{reason:'known-failure'}]:[]}));
  process.exit(0);
}
if (command.endsWith('-to-pdf')) {
  const input = process.argv[3];
  const output = process.argv[4];
  const sourceFormat = command.slice(0, -'-to-pdf'.length);
  const backend = ['doc','docx','odt','rtf'].includes(sourceFormat)?'wps':['xls','xlsx','ods'].includes(sourceFormat)?'et':'wpp';
  if (process.env.TIWATER_OFFICE_PDF_BACKEND !== backend) process.exit(3);
  const inputBytes = fs.readFileSync(input);
  const outputBytes = Buffer.from('%PDF-1.7 fake native render');
  fs.writeFileSync(output, outputBytes);
  const sha = bytes => crypto.createHash('sha256').update(bytes).digest('hex');
  const reportedOutputHash = input.includes('wrong-hash') ? '0'.repeat(64) : sha(outputBytes);
  process.stdout.write(JSON.stringify({status:'ok',input, input_sha256:sha(inputBytes),output,output_sha256:reportedOutputHash,source_format:sourceFormat,target_format:'pdf',version:'0.9.22',backend,fallback_reason:input.includes('fallback')?'non-native':null,page_count:1,native_render_provenance:{schema:'tiwater.convert-native-render-provenance/v1',backend,input:{sha256:sha(inputBytes),size_bytes:inputBytes.length},output:{sha256:reportedOutputHash,size_bytes:outputBytes.length},page_count:1}}));
  process.exit(0);
}
process.stderr.write('unexpected command: ' + command);
process.exit(2);
`, 'utf8');
  await chmod(fakeRuntime, 0o755);
  await symlink(fakeRuntime, path.join(temporary, 'tiwater-convert'));

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
    assert.equal(initialized.result.serverInfo.version, '0.10.0');
    const listed = await request('tools/list');
    const names = listed.result.tools.map(tool => tool.name);
    assert(names.includes('docx_list_migration_choices'));
    assert(names.includes('docx_query_migration_choices'));
    assert(names.includes('docx_migrate_template'));
    assert(names.includes('docx_verify_migration'));
    assert(names.includes('office_render_pdf'));
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
    assert(!names.includes('docx_validate_template_transform'));
    const migrateTool = listed.result.tools.find(tool => tool.name === 'docx_migrate_template');
    const listTool = listed.result.tools.find(tool => tool.name === 'docx_list_migration_choices');
    const queryTool = listed.result.tools.find(tool => tool.name === 'docx_query_migration_choices');
    assert.match(queryTool.description, /place-content moves current content/u);
    assert.match(queryTool.description, /keep-template-label preserves the target label and structure/u);
    assert.match(queryTool.description, /select-template-option marks a target option/u);
    assert.match(JSON.stringify(queryTool.inputSchema), /Choose the action from scenario meaning/u);
    assert.equal(listTool.inputSchema.required.includes('output'), true);
    assert.equal(listTool.outputSchema.properties.artifact.properties.sha256.type, 'string');
    assert.equal(listTool.outputSchema.properties.summary.properties.sourceCount.type, 'integer');
    assert.equal(migrateTool.inputSchema.required.includes('receiptOutput'), true);
    const migrationChoiceSchemas = migrateTool.inputSchema.properties.choices.items.anyOf;
    assert.equal(migrationChoiceSchemas.length, 2);
    assert.equal(migrationChoiceSchemas.some(schema => schema.required?.includes('alternativeRef')), true);
    assert.equal(migrationChoiceSchemas.every(schema => !('cardinality' in (schema.properties ?? {}))), true);
    assert.equal(migrationChoiceSchemas.every(schema => !('sourceChoiceId' in (schema.properties ?? {}))), true);
    assert.equal(migrationChoiceSchemas.every(schema => !('targetChoiceId' in (schema.properties ?? {}))), true);
    assert.equal(migrateTool.outputSchema.properties.artifact.properties.sha256.type, 'string');
    assert.equal(migrateTool.outputSchema.properties.summary.properties.outputVerified.type, 'boolean');
    const publicActions = migrationChoiceSchemas
      .flatMap(schema => schema.properties?.action?.enum ?? [])
      .sort();
    assert.deepEqual(publicActions, [
      'exclude-source',
      'review-source',
    ].sort());
    assert.equal(migrateTool.inputSchema.additionalProperties, false);

    const renderInput = path.join(temporary, 'render-current.docx');
    const renderOutput = path.join(temporary, 'rendered', 'current.pdf');
    const renderReceipt = path.join(temporary, 'rendered', 'current.receipt.json');
    await writeFile(renderInput, 'current document bytes', 'utf8');
    const rendered = await request('tools/call', {
      name: 'office_render_pdf',
      arguments: { input: renderInput, output: renderOutput, receiptOutput: renderReceipt },
    });
    assert.notEqual(rendered.result.isError, true, JSON.stringify(rendered.result));
    assert.equal(rendered.result.structuredContent.summary.backend, 'wps');
    assert.equal(rendered.result.structuredContent.summary.pageCount, 1);
    assert.equal(rendered.result.structuredContent.pdf.path, renderOutput);
    assert.equal(rendered.result.structuredContent.receipt.path, renderReceipt);
    assert.equal(JSON.parse(await readFile(renderReceipt, 'utf8')).backend, 'wps');
    for (const [extension, expectedBackend] of [['xlsx', 'et'], ['pptx', 'wpp']]) {
      const formatInput = path.join(temporary, `render-current.${extension}`);
      await writeFile(formatInput, `${extension} bytes`, 'utf8');
      const formatRender = await request('tools/call', {
        name: 'office_render_pdf',
        arguments: {
          input: formatInput,
          output: path.join(temporary, 'rendered', `${extension}.pdf`),
          receiptOutput: path.join(temporary, 'rendered', `${extension}.receipt.json`),
        },
      });
      assert.notEqual(formatRender.result.isError, true, JSON.stringify(formatRender.result));
      assert.equal(formatRender.result.structuredContent.summary.backend, expectedBackend);
    }
    const replayedRender = await request('tools/call', {
      name: 'office_render_pdf',
      arguments: { input: renderInput, output: renderOutput, receiptOutput: path.join(temporary, 'rendered', 'replay.receipt.json') },
    });
    assert.equal(replayedRender.result.isError, true);
    const unsupportedRender = await request('tools/call', {
      name: 'office_render_pdf',
      arguments: { input: path.join(temporary, 'current.txt'), output: path.join(temporary, 'rendered', 'invalid.pdf'), receiptOutput: path.join(temporary, 'rendered', 'invalid.receipt.json') },
    });
    assert.equal(unsupportedRender.result.isError, true);
    const nonPdfRender = await request('tools/call', {
      name: 'office_render_pdf',
      arguments: { input: renderInput, output: path.join(temporary, 'rendered', 'invalid.txt'), receiptOutput: path.join(temporary, 'rendered', 'invalid-output.receipt.json') },
    });
    assert.equal(nonPdfRender.result.isError, true);
    for (const invalidReceipt of ['fallback', 'wrong-hash']) {
      const invalidInput = path.join(temporary, `${invalidReceipt}.docx`);
      const invalidPdf = path.join(temporary, 'rendered', `${invalidReceipt}.pdf`);
      const invalidReceiptOutput = path.join(temporary, 'rendered', `${invalidReceipt}.receipt.json`);
      await writeFile(invalidInput, `${invalidReceipt} bytes`, 'utf8');
      const rejected = await request('tools/call', {
        name: 'office_render_pdf',
        arguments: { input: invalidInput, output: invalidPdf, receiptOutput: invalidReceiptOutput },
      });
      assert.equal(rejected.result.isError, true);
      await assert.rejects(readFile(invalidPdf));
      await assert.rejects(readFile(invalidReceiptOutput));
    }

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
    assert.match(sourcePage.result.structuredContent.items[0].ref, /^S1-[0-9a-f]{8}$/);
    const sourceRef = sourcePage.result.structuredContent.items[0].ref;
    assert.equal('id' in sourcePage.result.structuredContent.items[0], false);
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
        sourceRef,
        action: 'place-content',
        text: 'TARGET',
        offset: 0,
        limit: 10,
      },
    });
    assert.equal(targetPage.result.structuredContent.view, 'targets');
    assert.equal(targetPage.result.structuredContent.action, 'place-content');
    assert.equal(targetPage.result.structuredContent.source.ref, sourceRef);
    assert.equal(targetPage.result.structuredContent.items[0].action, 'place-content');
    assert.match(targetPage.result.structuredContent.items[0].ref, /^S1-P1-[0-9a-f]{8}$/);
    assert.match(targetPage.result.structuredContent.items[0].target.ref, /^T1-[0-9a-f]{8}$/);
    const alternativeRef = targetPage.result.structuredContent.items[0].ref;
    const targetRef = targetPage.result.structuredContent.items[0].target.ref;
    assert.deepEqual(targetPage.result.structuredContent.page, {
      offset: 0, returned: 1, total: 1, hasMore: false,
    });

    const allAlternatives = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'targets', sourceRef },
    });
    assert.deepEqual(allAlternatives.result.structuredContent.items.map(item => item.action), [
      'place-content', 'keep-template-label',
    ]);
    assert.equal(allAlternatives.result.structuredContent.source.allowedActions.includes('keep-template-content'), false);

    const hiddenDependentAction = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'targets', sourceRef, action: 'keep-template-content' },
    });
    assert.equal(hiddenDependentAction.result.isError, true);

    const unknownSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: sourceRef.replace(/^S1-/, 'S99-'), action: 'place-content' },
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
    assert.match(firstSource.result.structuredContent.items[0].ref, /^S1-[0-9a-f]{8}$/);
    const firstMultiSourceRef = firstSource.result.structuredContent.items[0].ref;
    assert.equal(firstSource.result.structuredContent.page.hasMore, true);
    const secondSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'sources', offset: 1, limit: 1 },
    });
    assert.match(secondSource.result.structuredContent.items[0].ref, /^S2-[0-9a-f]{8}$/);
    const secondMultiSourceRef = secondSource.result.structuredContent.items[0].ref;
    assert.equal(secondSource.result.structuredContent.page.hasMore, false);

    const contextMatch = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: {
        source: '/multi.docx',
        baseline: '/baseline.docx',
        view: 'targets',
        sourceRef: secondMultiSourceRef,
        action: 'place-content',
        text: 'REVISION',
      },
    });
    assert.deepEqual(contextMatch.result.structuredContent.items.map(item => item.target.ref.split('-')[0]), ['T2']);
    const revisionTargetRef = contextMatch.result.structuredContent.items[0].target.ref;
    const revisionAlternativeRef = contextMatch.result.structuredContent.items[0].ref;
    const alphaTarget = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: firstMultiSourceRef, action: 'place-content' },
    });
    const headerAlternativeRef = alphaTarget.result.structuredContent.items[0].ref;
    const betaTarget = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: secondMultiSourceRef, action: 'place-content', text: 'destination' },
    });
    const bodyAlternativeRef = betaTarget.result.structuredContent.items[0].ref;
    const contextSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/context.docx', baseline: '/baseline.docx', view: 'sources' },
    });
    const contextSourceRef = contextSource.result.structuredContent.items[0].ref;
    const structurallyRelated = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/context.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: contextSourceRef, action: 'place-content' },
    });
    assert.deepEqual(structurallyRelated.result.structuredContent.items.map(item => item.target.ref.split('-')[0]), ['T2', 'T1']);
    const headerSearch = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/context.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: contextSourceRef, action: 'place-content', text: 'Release identifier' },
    });
    assert.deepEqual(headerSearch.result.structuredContent.items.map(item => item.target.ref.split('-')[0]), ['T2']);
    const changedContextSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/context-mutated.docx', baseline: '/baseline.docx', view: 'sources' },
    });
    const changedContextChangesOrder = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/context-mutated.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: changedContextSource.result.structuredContent.items[0].ref, action: 'place-content' },
    });
    assert.deepEqual(changedContextChangesOrder.result.structuredContent.items.map(item => item.target.ref.split('-')[0]), ['T1', 'T2']);
    const noMatches = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: secondMultiSourceRef, action: 'place-content', text: 'not present' },
    });
    assert.deepEqual(noMatches.result.structuredContent.items, []);
    assert.equal(noMatches.result.structuredContent.page.total, 0);
    const opaqueIdIsNotVisibleText = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: secondMultiSourceRef, action: 'place-content', text: 'target-header' },
    });
    assert.deepEqual(opaqueIdIsNotVisibleText.result.structuredContent.items, []);

    const cleanupTargets = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/multi.docx', baseline: '/baseline.docx', view: 'cleanup', text: 'revision' },
    });
    assert.equal(cleanupTargets.result.structuredContent.source, null);
    assert.equal(cleanupTargets.result.structuredContent.action, null);
    assert.deepEqual(cleanupTargets.result.structuredContent.items.map(item => item.ref), [revisionTargetRef]);

    const manyCatalogPath = path.join(temporary, 'catalogs', 'many.json');
    await request('tools/call', {
      name: 'docx_list_migration_choices',
      arguments: { source: '/many.docx', baseline: '/baseline.docx', output: manyCatalogPath },
    });
    const manySource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/many.docx', baseline: '/baseline.docx', view: 'sources' },
    });
    const manySourceRef = manySource.result.structuredContent.items[0].ref;
    const boundedTargets = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/many.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: manySourceRef, action: 'place-content', limit: 10 },
    });
    assert.equal(boundedTargets.result.structuredContent.items.length, 10);
    assert.equal(boundedTargets.result.structuredContent.page.total, 25);
    assert.equal(boundedTargets.result.structuredContent.page.hasMore, true);
    const nextTargets = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/many.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: manySourceRef, action: 'place-content', offset: 10, limit: 10 },
    });
    const finalTargets = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/many.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: manySourceRef, action: 'place-content', offset: 20, limit: 10 },
    });
    assert.equal(finalTargets.result.structuredContent.page.hasMore, false);
    assert.equal(new Set([
      ...boundedTargets.result.structuredContent.items,
      ...nextTargets.result.structuredContent.items,
      ...finalTargets.result.structuredContent.items,
    ].map(item => item.target.ref)).size, 25);

    const invalidCatalogSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/invalid-output.docx', baseline: '/baseline.docx', view: 'sources' },
    });
    assert.equal(invalidCatalogSource.result.isError, true);

    const actionNotAllowed = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/current.docx', baseline: '/baseline.docx', view: 'targets', sourceRef, action: 'select-template-option' },
    });
    assert.equal(actionNotAllowed.result.isError, true);

    const driftSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/drift.docx', baseline: '/baseline.docx', view: 'sources' },
    });
    const targetDrift = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/drift.docx', baseline: '/baseline.docx', view: 'targets', sourceRef: driftSource.result.structuredContent.items[0].ref, action: 'place-content' },
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

    const referencedChoices = [{ alternativeRef }];
    const choices = [{ sourceChoiceId: 'source-1', action: 'place-content', targetChoiceId: 'target-1' }];
    const migrationReceiptPath = path.join(temporary, 'receipts', 'migration.json');
    const migrated = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/output.docx',
        receiptOutput: migrationReceiptPath, choices: referencedChoices,
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

    const legacyIdentityReceiptPath = path.join(temporary, 'receipts', 'legacy-identity.json');
    const legacyIdentity = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/legacy-identity.docx',
        receiptOutput: legacyIdentityReceiptPath,
        choices: [{ sourceChoiceId: 'source-1', action: 'keep-template-content', targetChoiceId: 'target-1' }],
      },
    });
    assert.equal(legacyIdentity.result.isError, true);
    await assert.rejects(readFile(legacyIdentityReceiptPath, 'utf8'));

    const referencedCleanupReceiptPath = path.join(temporary, 'receipts', 'referenced-cleanup.json');
    const referencedCleanup = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/multi.docx', baseline: '/baseline.docx', output: '/multi-output.docx',
        receiptOutput: referencedCleanupReceiptPath,
        choices: [
          { alternativeRef: headerAlternativeRef },
          { alternativeRef: bodyAlternativeRef },
        ],
        templateCleanup: [{ targetRef: revisionTargetRef, scope: 'row' }],
      },
    });
    assert.equal(referencedCleanup.result.structuredContent.summary.pass, true);
    assert.deepEqual(JSON.parse(await readFile(referencedCleanupReceiptPath, 'utf8')).payload, {
      schema: 'tiwater.docx.template-migration-business-choices/v1',
      choices: [
        { sourceChoiceId: 'source-alpha', action: 'place-content', targetChoiceId: 'target-header' },
        { sourceChoiceId: 'source-beta', action: 'place-content', targetChoiceId: 'target-body' },
      ],
      templateCleanup: [{ targetChoiceId: 'target-revision', scope: 'row' }],
    });

    const claimedCleanupReceiptPath = path.join(temporary, 'receipts', 'claimed-cleanup.json');
    const claimedCleanup = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/multi.docx', baseline: '/baseline.docx', output: '/claimed-cleanup.docx',
        receiptOutput: claimedCleanupReceiptPath,
        choices: [
          { alternativeRef: headerAlternativeRef },
          { alternativeRef: revisionAlternativeRef },
        ],
        templateCleanup: [{ targetRef: revisionTargetRef, scope: 'cell' }],
      },
    });
    assert.equal(claimedCleanup.result.structuredContent.summary.pass, true);
    assert.deepEqual(JSON.parse(await readFile(claimedCleanupReceiptPath, 'utf8')).payload, {
      schema: 'tiwater.docx.template-migration-business-choices/v1',
      choices: [
        { sourceChoiceId: 'source-alpha', action: 'place-content', targetChoiceId: 'target-header' },
        { sourceChoiceId: 'source-beta', action: 'place-content', targetChoiceId: 'target-revision' },
      ],
    });

    const invalidReferenceReceipt = path.join(temporary, 'receipts', 'invalid-reference.json');
    const invalidReference = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/invalid-reference.docx',
        receiptOutput: invalidReferenceReceipt,
        choices: [{ alternativeRef: alternativeRef.replace(/-P1-/, '-P99-') }],
      },
    });
    assert.equal(invalidReference.result.isError, true);
    await assert.rejects(readFile(invalidReferenceReceipt, 'utf8'));

    const staleReferenceReceipt = path.join(temporary, 'receipts', 'stale-reference.json');
    const staleReference = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/changed.docx', baseline: '/baseline.docx', output: '/stale-reference.docx',
        receiptOutput: staleReferenceReceipt,
        choices: referencedChoices,
      },
    });
    assert.equal(staleReference.result.isError, true);
    assert.match(staleReference.result.content[0].text, /stale or invalid migration alternative ref/);
    await assert.rejects(readFile(staleReferenceReceipt, 'utf8'));

    const recombinedReferenceReceipt = path.join(temporary, 'receipts', 'recombined-reference.json');
    const recombinedReference = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/recombined-reference.docx',
        receiptOutput: recombinedReferenceReceipt,
        choices: [{ alternativeRef: alternativeRef.replace('-P1-', '-K1-') }],
      },
    });
    assert.equal(recombinedReference.result.isError, true);
    assert.match(recombinedReference.result.content[0].text, /Invalid string/);
    await assert.rejects(readFile(recombinedReferenceReceipt, 'utf8'));

    const mixedIdentity = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/mixed.docx',
        receiptOutput: path.join(temporary, 'receipts', 'mixed.json'),
        choices: [{ alternativeRef, sourceRef }],
      },
    });
    assert.equal(mixedIdentity.result.isError, true);

    const missingTarget = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/missing-target.docx',
        receiptOutput: path.join(temporary, 'receipts', 'missing-target.json'),
        choices: [{ sourceRef, action: 'place-content' }],
      },
    });
    assert.equal(missingTarget.result.isError, true);

    const mixedTarget = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/mixed-target.docx',
        receiptOutput: path.join(temporary, 'receipts', 'mixed-target.json'),
        choices: [{ sourceRef, action: 'place-content', targetRef, targetChoiceId: 'target-1' }],
      },
    });
    assert.equal(mixedTarget.result.isError, true);

    const terminalWithTarget = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/all.docx', baseline: '/baseline.docx', output: '/terminal-target.docx',
        receiptOutput: path.join(temporary, 'receipts', 'terminal-target.json'),
        choices: [{ sourceRef, action: 'exclude-source', targetRef }],
      },
    });
    assert.equal(terminalWithTarget.result.isError, true);

    const allReceiptPath = path.join(temporary, 'receipts', 'all.json');
    const allSource = await request('tools/call', {
      name: 'docx_query_migration_choices',
      arguments: { source: '/all.docx', baseline: '/baseline.docx', view: 'sources' },
    });
    const all = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/all.docx', baseline: '/baseline.docx', output: '/all-output.docx',
        receiptOutput: allReceiptPath,
        choices: [{ sourceRef: allSource.result.structuredContent.items[0].ref, action: 'exclude-source' }],
      },
    });
    assert.equal(all.result.structuredContent.summary.pass, true);
    assert.deepEqual(JSON.parse(await readFile(allReceiptPath, 'utf8')).payload.choices, [
      { sourceChoiceId: 'source-all', action: 'exclude-source', cardinality: 'all' },
    ]);

    const incomplete = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/multi.docx', baseline: '/baseline.docx', output: '/incomplete.docx',
        receiptOutput: path.join(temporary, 'receipts', 'incomplete.json'),
        choices: [{ alternativeRef: headerAlternativeRef }],
      },
    });
    assert.equal(incomplete.result.isError, true);
    assert.match(incomplete.result.content[0].text, /migration choices must cover every source id/);

    const silent = await request('tools/call', {
      name: 'docx_migrate_template',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/silent.docx',
        receiptOutput: path.join(temporary, 'receipts', 'silent.json'), choices: referencedChoices,
      },
    });
    assert.equal(silent.result.isError, true);
    assert.match(silent.result.content[0].text, /typed-migration-refusal/);

    const verificationReceiptPath = path.join(temporary, 'receipts', 'verification.json');
    const verified = await request('tools/call', {
      name: 'docx_verify_migration',
      arguments: {
        source: '/current.docx', baseline: '/baseline.docx', output: '/output.docx',
        receiptOutput: verificationReceiptPath, choices: referencedChoices,
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
        receiptOutput: failedReceiptPath, choices: referencedChoices,
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
