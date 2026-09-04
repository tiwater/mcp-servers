import assert from 'node:assert/strict';
import { readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { resolveFileBackedChanges } from './file-backed-changes.mjs';

const schemaPath = new URL('../../packages/docx-cli/contracts/mcp-input/docx_replace_content_from_source.schema.json', import.meta.url);

test('published contract exposes independent inline and file-backed inputs', async () => {
  const schema = JSON.parse(await readFile(schemaPath, 'utf8'));
  assert.equal(schema.properties.changes.type, 'array');
  assert.equal(schema.properties.changes.minItems, 1);
  assert.equal(schema.properties.changesInput.type, 'string');
  assert.equal(schema.properties.changesInput['x-tiwater-file-role'], 'read');
  assert.deepEqual(schema.required, ['input', 'output', 'receiptOutput']);
  assert.equal('anyOf' in schema, false);
});

test('leaves a small inline changes array unchanged', async () => {
  const args = { changes: [{ value: 'inline' }] };
  assert.equal(await resolveFileBackedChanges(args), args);
});

test('loads a large changes array from its file without changing its items', async () => {
  const file = path.join(os.tmpdir(), `tiwater-changes-${process.pid}-${Date.now()}.json`);
  const changes = Array.from({ length: 512 }, (_, index) => ({ index }));
  await writeFile(file, JSON.stringify(changes));
  const resolved = await resolveFileBackedChanges({ input: 'a.docx', changesInput: file, output: 'b.docx' });
  assert.deepEqual(resolved.changes, changes);
  assert.equal('changesInput' in resolved, false);
});

test('appends file-backed changes after inline changes in one atomic request', async () => {
  const file = path.join(os.tmpdir(), `tiwater-combined-changes-${process.pid}-${Date.now()}.json`);
  await writeFile(file, JSON.stringify([{ value: 'file' }]));
  const resolved = await resolveFileBackedChanges({ changes: [{ value: 'inline' }], changesInput: file });
  assert.deepEqual(resolved.changes, [{ value: 'inline' }, { value: 'file' }]);
  assert.equal('changesInput' in resolved, false);
});

test('rejects missing, malformed, empty, and non-array inputs', async () => {
  await assert.rejects(resolveFileBackedChanges({}), /provide-changes-or-changesInput/);
  const base = path.join(os.tmpdir(), `tiwater-invalid-changes-${process.pid}-${Date.now()}`);
  await writeFile(`${base}-syntax.json`, '{');
  await writeFile(`${base}-empty.json`, '[]');
  await writeFile(`${base}-object.json`, '{}');
  await assert.rejects(resolveFileBackedChanges({ changesInput: `${base}-syntax.json` }), /changesInput-invalid/);
  await assert.rejects(resolveFileBackedChanges({ changesInput: `${base}-empty.json` }), /non-empty-json-array/);
  await assert.rejects(resolveFileBackedChanges({ changesInput: `${base}-object.json` }), /non-empty-json-array/);
});
