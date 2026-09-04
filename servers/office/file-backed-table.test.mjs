import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { resolveFileBackedTable } from './file-backed-table.mjs';

const schemaPath = new URL('../../packages/docx-cli/contracts/mcp-input/docx_set_table.schema.json', import.meta.url);

test('published contract accepts inline fields or a file-backed table input', async () => {
  const schema = JSON.parse(await readFile(schemaPath, 'utf8'));
  assert.equal(schema.properties.tableInput.type, 'string');
  assert.equal(schema.properties.tableInput['x-tiwater-file-role'], 'read');
  assert.equal(schema.properties.rows.items.properties.cells.minItems, 0);
  assert.deepEqual(schema.required, ['input', 'output', 'receiptOutput']);
});

test('leaves a complete inline table request unchanged', async () => {
  const args = { table: {}, existingRows: {}, columns: [], rows: [] };
  assert.equal(await resolveFileBackedTable(args), args);
});

test('loads an exact table request from disk', async () => {
  const file = path.join(os.tmpdir(), `tiwater-table-${process.pid}-${Date.now()}.json`);
  const tableRequest = { table: { path: '/table' }, existingRows: { first: {}, last: {} }, columns: [{ id: 'a' }], rows: [] };
  await writeFile(file, JSON.stringify(tableRequest));
  const resolved = await resolveFileBackedTable({ input: 'a.docx', tableInput: file, output: 'b.docx' });
  assert.deepEqual(resolved, { input: 'a.docx', output: 'b.docx', ...tableRequest });
});

test('rejects missing, duplicate, malformed, and incomplete representations', async () => {
  await assert.rejects(resolveFileBackedTable({}), /provide-inline-table-or-tableInput/);
  await assert.rejects(resolveFileBackedTable({ table: {}, existingRows: {}, columns: [], rows: [], tableInput: 'x' }), /provide-only-one-table-input/);
  const base = path.join(os.tmpdir(), `tiwater-invalid-table-${process.pid}-${Date.now()}`);
  await writeFile(`${base}-syntax.json`, '{');
  await writeFile(`${base}-incomplete.json`, '{}');
  await assert.rejects(resolveFileBackedTable({ tableInput: `${base}-syntax.json` }), /tableInput-invalid/);
  await assert.rejects(resolveFileBackedTable({ tableInput: `${base}-incomplete.json` }), /table-existingRows-columns-and-rows/);
});
