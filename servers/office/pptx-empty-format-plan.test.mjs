import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

const providerSchema = new URL('../../packages/pptx-cli/contracts/mcp-input/pptx_apply_format.schema.json', import.meta.url);
const publishedSchema = new URL('./contracts/pptx_apply_format.schema.json', import.meta.url);

test('published PPTX format contract accepts an empty materialization plan', async () => {
  const [providerBytes, publishedBytes] = await Promise.all([
    readFile(providerSchema),
    readFile(publishedSchema),
  ]);
  assert.deepEqual(publishedBytes, providerBytes);

  const schema = JSON.parse(providerBytes);
  assert.equal(schema.properties.changes.type, 'array');
  assert.equal(schema.properties.changes.minItems, 0);
  assert.match(schema.properties.changes.description, /empty.*unchanged output.*passing receipt/iu);
  assert.ok(schema.required.includes('changes'));
});
