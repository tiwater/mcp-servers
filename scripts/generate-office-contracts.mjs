#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, readdir, unlink, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const serverRoot = path.join(repoRoot, 'servers');
const outputRoot = path.join(serverRoot, 'office', 'contracts');
const manifestName = 'tiwater-office-provider-contract-manifest-v1.json';
const providerRoots = [
  'packages/convert-cli/schemas/mcp-input',
  'packages/docx-cli/contracts/mcp-input',
  'packages/pptx-cli/contracts/mcp-input',
  'packages/xlsx-cli/contracts/mcp-input',
];

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

await mkdir(outputRoot, { recursive: true });
const contracts = [];
for (const relativeRoot of providerRoots) {
  const absoluteRoot = path.join(repoRoot, relativeRoot);
  for (const name of (await readdir(absoluteRoot)).filter(name => name.endsWith('.schema.json')).sort()) {
    const toolName = name.slice(0, -'.schema.json'.length);
    const source = path.join(absoluteRoot, name);
    const bytes = await readFile(source);
    JSON.parse(bytes.toString('utf8'));
    contracts.push({ toolName, source, sourceRelative: `${relativeRoot}/${name}`, name, hash: sha256(bytes) });
  }
}

const duplicateNames = contracts.filter((contract, index) =>
  contracts.findIndex(candidate => candidate.toolName === contract.toolName) !== index);
if (duplicateNames.length > 0) {
  throw new Error(`Duplicate provider-owned MCP contracts: ${duplicateNames.map(item => item.toolName).join(', ')}`);
}

const expectedFiles = new Set(contracts.map(contract => contract.name));
for (const name of await readdir(outputRoot)) {
  if (name.endsWith('.schema.json') && !expectedFiles.has(name)) {
    await unlink(path.join(outputRoot, name));
  }
}
for (const contract of contracts) {
  await copyFile(contract.source, path.join(outputRoot, contract.name));
}

const packageJson = JSON.parse(await readFile(path.join(serverRoot, 'package.json'), 'utf8'));
const manifest = {
  schema: 'tiwater.office-provider-contract-manifest/v1',
  provider: { id: packageJson.name, version: packageJson.version },
  tools: contracts.sort((left, right) => left.toolName.localeCompare(right.toolName)).map(contract => ({
    name: contract.toolName,
    providerContract: { source: contract.sourceRelative, sha256: contract.hash },
    inputContract: { path: `office/contracts/${contract.name}`, sha256: contract.hash },
  })),
};
await writeFile(path.join(outputRoot, manifestName), `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`Generated ${contracts.length} Office MCP input contracts from provider-owned schemas.`);
