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
const crypto = require('node:crypto');
const path = require('node:path');
const command = process.argv[2];
if (command === 'edit') {
  const input = process.argv[3];
  const operations = JSON.parse(fs.readFileSync(process.argv[4], 'utf8')).operations;
  const output = process.argv[5];
  fs.copyFileSync(input, output);
  const docx = path.basename(process.argv[1]) === 'tiwater-docx';
  process.stdout.write(JSON.stringify(docx ? {
    Input: input, Output: output,
    AppliedOperations: operations.map(operation => ({Type: operation.type, Applied: true, Detail: 'ok'})),
  } : {
    input, output,
    appliedOperations: operations.map(operation => ({type: operation.type, applied: true, detail: 'ok'})),
  }));
  process.exit(0);
}
if (command === 'set-shape-geometry' || command === 'replace-picture-image') {
  const input = process.argv[3];
  const changes = JSON.parse(fs.readFileSync(process.argv[4], 'utf8')).changes;
  const output = process.argv[5];
  fs.copyFileSync(input, output);
  const geometry = command === 'set-shape-geometry';
  const rejected = geometry && changes.some(change => change.x === -999);
  process.stdout.write(JSON.stringify({
    input, output, operationCount: changes.length, appliedCount: rejected ? 0 : changes.length,
    issues: rejected ? [{slideNumber: changes[0].slideNumber, shapeId: changes[0].shapeId, message: 'unsupported test mutation'}] : [],
    changes: rejected ? [] : changes.map(change => geometry
      ? {slideNumber: change.slideNumber, shapeId: change.shapeId, before: {x: 1, y: 2, cx: 3, cy: 4}, after: {x: change.x === -998 ? -997 : change.x, y: change.y, cx: change.cx, cy: change.cy}}
      : {slideNumber: change.slideNumber, shapeId: change.shapeId, image: change.image, beforeSha256: '0'.repeat(64), afterSha256: crypto.createHash('sha256').update(fs.readFileSync(change.image)).digest('hex')}),
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
    assert.equal(initialized.result.serverInfo.version, '0.14.3');

    const listed = await request('tools/list');
    const names = listed.result.tools.map(tool => tool.name);
    for (const removed of ['docx_list_migration_choices', 'docx_query_migration_choices', 'docx_migrate_template', 'docx_verify_migration', 'xlsx_apply']) {
      assert(!names.includes(removed), `${removed} must not be public`);
    }
    for (const required of [
      'docx_inspect', 'docx_inspect_tables', 'docx_set_table_cell_text', 'docx_validate',
      'docx_insert_body_range', 'docx_replace_drawing_image', 'docx_insert_body_image',
      'docx_set_table_row_repeat_as_header',
      'xlsx_convert_legacy', 'xlsx_set_cell_value', 'xlsx_delete_rows', 'xlsx_set_page_setup', 'xlsx_validate',
      'pptx_inspect', 'pptx_apply_template', 'pptx_apply_format', 'pptx_set_shape_geometry', 'pptx_replace_picture_image', 'pptx_validate',
      'office_render_pdf',
    ]) assert(names.includes(required), `missing ${required}`);

    assert(!names.some(name => /scenario|migration|issue|workitem/i.test(name)));

    const tool = listed.result.tools.find(candidate => candidate.name === 'docx_set_table_cell_text');
    assert.deepEqual(tool.inputSchema.required.sort(), ['changes', 'input', 'output', 'receiptOutput']);
    assert.equal(tool.inputSchema.properties.changes.items.properties.type, undefined);
    assert.equal(tool.inputSchema.properties.changes.items.additionalProperties, false);

    const deleteRowsTool = listed.result.tools.find(candidate => candidate.name === 'xlsx_delete_rows');
    assert.deepEqual(deleteRowsTool.inputSchema.properties.changes.items.required.sort(), ['count', 'sheet', 'startRow']);
    assert.deepEqual(Object.keys(deleteRowsTool.inputSchema.properties.changes.items.properties).sort(), ['count', 'sheet', 'startRow']);
    assert.equal(deleteRowsTool.inputSchema.properties.changes.items.additionalProperties, false);

    const repeatHeaderTool = listed.result.tools.find(candidate => candidate.name === 'docx_set_table_row_repeat_as_header');
    assert.deepEqual(repeatHeaderTool.inputSchema.required.sort(), ['changes', 'input', 'output', 'receiptOutput']);
    assert.equal(repeatHeaderTool.inputSchema.properties.changes.items.anyOf.length, 3);
    const repeatShapes = repeatHeaderTool.inputSchema.properties.changes.items.anyOf
      .map(shape => Object.keys(shape.properties).sort());
    assert.deepEqual(repeatShapes, [
      ['repeatAsHeader', 'rowIndex', 'tableIndex'],
      ['headerIndex', 'repeatAsHeader', 'rowIndex', 'tableIndex'],
      ['footerIndex', 'repeatAsHeader', 'rowIndex', 'tableIndex'],
    ]);
    assert(repeatHeaderTool.inputSchema.properties.changes.items.anyOf.every(shape => shape.additionalProperties === false));

    const formatTool = listed.result.tools.find(candidate => candidate.name === 'pptx_apply_format');
    assert.equal(formatTool.inputSchema.properties.changes.items.properties.x, undefined);
    assert.equal(formatTool.inputSchema.properties.changes.items.properties.image, undefined);
    const geometryTool = listed.result.tools.find(candidate => candidate.name === 'pptx_set_shape_geometry');
    assert.deepEqual(geometryTool.inputSchema.required.sort(), ['changes', 'input', 'output', 'receiptOutput']);
    assert.deepEqual(Object.keys(geometryTool.inputSchema.properties.changes.items.properties).sort(), ['cx', 'cy', 'shapeId', 'slideNumber', 'x', 'y']);
    assert.equal(geometryTool.inputSchema.properties.changes.items.properties.type, undefined);
    assert.equal(geometryTool.inputSchema.properties.changes.items.additionalProperties, false);

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

    const repeatedOutput = path.join(temporary, 'repeat-header.docx');
    const repeatedReceipt = path.join(temporary, 'repeat-header.json');
    const repeated = await request('tools/call', {
      name: 'docx_set_table_row_repeat_as_header',
      arguments: { input, output: repeatedOutput, receiptOutput: repeatedReceipt, changes: [{ headerIndex: 0, tableIndex: 1, rowIndex: 2, repeatAsHeader: false }] },
    });
    assert.equal(repeated.result.structuredContent.summary.pass, true);
    const repeatReceipt = JSON.parse(await readFile(repeatedReceipt, 'utf8'));
    assert.equal(repeatReceipt.operationType, 'setTableRowRepeatAsHeader');
    assert.deepEqual(repeatReceipt.appliedOperations.map(operation => operation.type), ['setTableRowRepeatAsHeader']);

    const repeatInjected = await request('tools/call', {
      name: 'docx_set_table_row_repeat_as_header',
      arguments: { input, output: path.join(temporary, 'bad-repeat.docx'), receiptOutput: path.join(temporary, 'bad-repeat.json'), changes: [{ tableIndex: 0, rowIndex: 0, repeatAsHeader: true, type: 'deleteTableRows' }] },
    });
    assert.equal(repeatInjected.result.isError, true);

    const workbookInput = path.join(temporary, 'input.xlsx');
    const workbookOutput = path.join(temporary, 'output.xlsx');
    const workbookReceiptOutput = path.join(temporary, 'xlsx-receipt.json');
    await writeFile(workbookInput, 'current workbook', 'utf8');
    const deletedRows = await request('tools/call', {
      name: 'xlsx_delete_rows',
      arguments: { input: workbookInput, output: workbookOutput, receiptOutput: workbookReceiptOutput, changes: [{ sheet: 'Data', startRow: 3, count: 2 }] },
    });
    assert.equal(deletedRows.result.structuredContent.summary.pass, true);
    const workbookReceipt = JSON.parse(await readFile(workbookReceiptOutput, 'utf8'));
    assert.equal(workbookReceipt.operationType, 'deleteRows');
    assert.deepEqual(workbookReceipt.appliedOperations.map(operation => operation.type), ['deleteRows']);

    const deleteRowsInjected = await request('tools/call', {
      name: 'xlsx_delete_rows',
      arguments: { input: workbookInput, output: path.join(temporary, 'injected.xlsx'), receiptOutput: path.join(temporary, 'injected.json'), changes: [{ sheet: 'Data', startRow: 3, count: 2, type: 'setCellValue' }] },
    });
    assert.equal(deleteRowsInjected.result.isError, true);

    const pptxInput = path.join(temporary, 'input.pptx');
    await writeFile(pptxInput, 'current presentation', 'utf8');
    const geometryOutput = path.join(temporary, 'geometry.pptx');
    const geometryReceipt = path.join(temporary, 'geometry.json');
    const geometry = await request('tools/call', {
      name: 'pptx_set_shape_geometry',
      arguments: { input: pptxInput, output: geometryOutput, receiptOutput: geometryReceipt, changes: [{ slideNumber: 1, shapeId: 2, x: -10, y: 20, cx: 30, cy: 40 }] },
    });
    assert.equal(geometry.result.structuredContent.summary.pass, true);
    assert.equal(geometry.result.structuredContent.summary.appliedCount, 1);
    assert.equal(JSON.parse(await readFile(geometryReceipt, 'utf8')).tool, 'pptx_set_shape_geometry');

    const rejectedOutput = path.join(temporary, 'rejected-geometry.pptx');
    const rejectedGeometry = await request('tools/call', {
      name: 'pptx_set_shape_geometry',
      arguments: { input: pptxInput, output: rejectedOutput, receiptOutput: path.join(temporary, 'rejected-geometry.json'), changes: [{ slideNumber: 1, shapeId: 2, x: -999, y: 20, cx: 30, cy: 40 }] },
    });
    assert.equal(rejectedGeometry.result.structuredContent.summary.pass, false);
    await assert.rejects(readFile(rejectedOutput));

    const mutatedOutput = path.join(temporary, 'mutated-geometry.pptx');
    const mutatedGeometry = await request('tools/call', {
      name: 'pptx_set_shape_geometry',
      arguments: { input: pptxInput, output: mutatedOutput, receiptOutput: path.join(temporary, 'mutated-geometry.json'), changes: [{ slideNumber: 1, shapeId: 2, x: -998, y: 20, cx: 30, cy: 40 }] },
    });
    assert.equal(mutatedGeometry.result.structuredContent.summary.pass, false);
    await assert.rejects(readFile(mutatedOutput));

    const image = path.join(temporary, 'replacement.png');
    await writeFile(image, Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    const picture = await request('tools/call', {
      name: 'pptx_replace_picture_image',
      arguments: { input: pptxInput, output: path.join(temporary, 'picture.pptx'), receiptOutput: path.join(temporary, 'picture.json'), changes: [{ slideNumber: 1, shapeId: 3, image }] },
    });
    assert.equal(picture.result.structuredContent.summary.pass, true);
    assert.equal(picture.result.structuredContent.summary.appliedCount, 1);

    const source = path.join(temporary, 'source.docx');
    await writeFile(source, 'source document', 'utf8');
    const insertedOutput = path.join(temporary, 'inserted.docx');
    const insertedReceipt = path.join(temporary, 'inserted.json');
    const inserted = await request('tools/call', {
      name: 'docx_insert_body_range',
      arguments: { input, output: insertedOutput, receiptOutput: insertedReceipt, changes: [{ source, sourceStartBodyIndex: 0, sourceEndBodyIndex: 1, targetBodyIndex: 0 }] },
    });
    assert.equal(inserted.result.structuredContent.summary.pass, true);
    const sourceBoundReceipt = JSON.parse(await readFile(insertedReceipt, 'utf8'));
    assert.equal(sourceBoundReceipt.operationType, 'insertBodyRange');
    assert.equal(sourceBoundReceipt.sources.length, 1);
    assert.equal(sourceBoundReceipt.sources[0].path, source);
    assert.match(sourceBoundReceipt.sources[0].sha256, /^[0-9a-f]{64}$/);
    assert.equal(sourceBoundReceipt.sourceBindingStable, true);
    const rangeTool = listed.result.tools.find(candidate => candidate.name === 'docx_insert_body_range');
    assert.deepEqual(rangeTool.inputSchema.properties.changes.items.required.sort(), ['source', 'sourceEndBodyIndex', 'sourceStartBodyIndex', 'targetBodyIndex']);
    assert.equal(rangeTool.inputSchema.properties.changes.items.properties.scenarioId, undefined);
    assert.equal(rangeTool.inputSchema.properties.changes.items.properties.type, undefined);

    const injected = await request('tools/call', {
      name: 'docx_set_table_cell_text',
      arguments: { input, output: path.join(temporary, 'bad.docx'), receiptOutput: path.join(temporary, 'bad.json'), changes: [{ tableIndex: 0, rowIndex: 0, cellIndex: 0, text: 'x', type: 'deleteBodyRange' }] },
    });
    assert.equal(injected.result.isError, true);

    const geometryInjected = await request('tools/call', {
      name: 'pptx_set_shape_geometry',
      arguments: { input: pptxInput, output: path.join(temporary, 'bad-geometry.pptx'), receiptOutput: path.join(temporary, 'bad-geometry.json'), changes: [{ slideNumber: 1, shapeId: 2, x: 1, y: 2, cx: 3, cy: 4, type: 'replace-picture-image' }] },
    });
    assert.equal(geometryInjected.result.isError, true);
  } finally {
    if (child.exitCode === null) {
      child.kill('SIGTERM');
      await new Promise(resolve => child.once('exit', resolve));
    }
    await rm(temporary, { recursive: true, force: true });
  }
});
