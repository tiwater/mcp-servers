#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const distributionRoot = path.join(repoRoot, 'servers');
const contractsRoot = path.join(distributionRoot, 'pdf', 'contracts');
const packageJson = JSON.parse(await readFile(path.join(distributionRoot, 'package.json'), 'utf8'));

const commonProperties = {
  input: {
    type: 'string',
    minLength: 1,
    description: 'Path to the current PDF revision.',
    'x-tiwater-file-role': 'read',
  },
  returnContent: {
    type: 'boolean',
    description: 'Return the complete result directly when it fits the published response limit.',
  },
  output: {
    type: 'string',
    minLength: 1,
    description: 'Absolute path for the immutable JSON observation artifact. An identical replay may reuse identical bytes; different content is never written over it.',
    'x-tiwater-file-role': 'write',
    'x-tiwater-file-effect': false,
  },
};

function contract(properties = {}, required = ['input', 'returnContent']) {
  return {
    '$schema': 'https://json-schema.org/draft/2020-12/schema',
    type: 'object',
    properties: { ...commonProperties, ...properties },
    required,
    additionalProperties: false,
  };
}

const contracts = new Map([
  ['pdf_inspect', contract({}, ['input', 'output'])],
  ['pdf_extract_tables', contract({
    pages: { type: 'array', minItems: 1, uniqueItems: true, items: { type: 'integer', minimum: 1 } },
    autoSpan: { type: 'boolean' },
  })],
  ['pdf_find_table', contract({
    name: { type: 'string', minLength: 1 },
    autoSpan: { type: 'boolean' },
  }, ['input', 'name', 'returnContent'])],
  ['pdf_ocr', contract({
    pages: { type: 'array', minItems: 1, uniqueItems: true, items: { type: 'integer', minimum: 1 } },
  })],
  ['pdf_extract_table_details', contract({
    pages: { type: 'array', minItems: 1, uniqueItems: true, items: { type: 'integer', minimum: 1 } },
  })],
]);

await mkdir(contractsRoot, { recursive: true });
const tools = [];
for (const [name, schema] of contracts) {
  const filename = `${name}.schema.json`;
  const bytes = `${JSON.stringify(schema, null, 2)}\n`;
  await writeFile(path.join(contractsRoot, filename), bytes);
  tools.push({
    name,
    inputContract: {
      path: `contracts/${filename}`,
      sha256: createHash('sha256').update(bytes).digest('hex'),
    },
  });
}

await writeFile(path.join(contractsRoot, 'tiwater-pdf-provider-contract-manifest-v1.json'), `${JSON.stringify({
  schema: 'tiwater.pdf-provider-contract-manifest/v1',
  provider: { id: packageJson.name, version: packageJson.version },
  runtime: { command: 'tiwater-pdf' },
  tools,
}, null, 2)}\n`);
