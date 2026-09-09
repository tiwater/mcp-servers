import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import * as z from 'zod/v4';

const schemaPath = new URL('../../packages/docx-cli/contracts/mcp-input/docx_read_all_tables.schema.json', import.meta.url);
const generatedSchemaPath = new URL('./contracts/docx_read_all_tables.schema.json', import.meta.url);
const schemaText = await readFile(schemaPath, 'utf8');
const schema = JSON.parse(schemaText);
const contract = z.fromJSONSchema(schema);

test('accepts the complete file-backed request with or without the false compatibility flag', () => {
  const required = { input: '/current/source.docx', output: '/current/tables.json' };
  assert.deepEqual(contract.parse(required), required);
  assert.deepEqual(contract.parse({ ...required, returnContent: false }), { ...required, returnContent: false });
});

test('publishes the exact provider-owned closed contract', async () => {
  assert.equal(await readFile(generatedSchemaPath, 'utf8'), schemaText);
  assert.deepEqual(schema.required, ['input', 'output']);
  assert.equal(schema.additionalProperties, false);
  assert.equal(schema.properties.returnContent.const, false);
});

test('rejects inline content, incomplete requests, and unrelated arguments', () => {
  const required = { input: '/unseen/source.docx', output: '/unseen/tables.json' };
  assert.throws(() => contract.parse({ ...required, returnContent: true }));
  assert.throws(() => contract.parse({ input: required.input, returnContent: false }));
  assert.throws(() => contract.parse({ ...required, page: 1 }));
});
