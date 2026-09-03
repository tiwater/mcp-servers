#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { constants as fsConstants } from 'node:fs';
import { access, mkdir, mkdtemp, readFile, readdir, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { promisify } from 'node:util';
import { writeIdempotentJsonArtifact } from '../servers/_shared/large-json-result.mjs';
import {
  assertEvidenceToolContract,
  evidenceRoleMetadataKey,
} from '../servers/_shared/evidence-role.mjs';
import {
  assertEffectKindToolContract,
  effectKindMetadataKey,
} from '../servers/_shared/effect-kind.mjs';

const execFileAsync = promisify(execFile);
const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');
const serverRoot = path.join(repoRoot, 'servers');
const serverPackagePath = path.join(serverRoot, 'package.json');
const serverLockPath = path.join(serverRoot, 'package-lock.json');
const generatedManifestRelativePath = 'office/contracts/tiwater-office-provider-contract-manifest-v1.json';
const textManifestRelativePath = 'text/contracts/tiwater-text-provider-contract-manifest-v1.json';
const generatedContractDeclaration = 'office/contracts/*.schema.json';
const textContractDeclaration = 'text/contracts/*.schema.json';
const fileRoleKey = 'x-tiwater-file-role';
const fileEffectKey = 'x-tiwater-file-effect';
const nativeDocxPathPattern = '^\\/[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*\\[[1-9][0-9]*\\](?:\\/[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*\\[[1-9][0-9]*\\])*$';
const expectedFileRolesByProperty = new Map([
  ['input', 'read'],
  ['baseline', 'read'],
  ['updated', 'read'],
  ['template', 'read'],
  ['image', 'read'],
  ['source', 'read'],
  ['output', 'write'],
  ['receiptOutput', 'write'],
]);
const expectedFileEffectsByProperty = new Map([
  ['receiptOutput', false],
]);
const requiredPackageFiles = [
  'package.json',
  '_shared/tool-runtime.mjs',
  '_shared/evidence-role.mjs',
  '_shared/effect-kind.mjs',
  'office/index.mjs',
  'office/README.md',
  generatedManifestRelativePath,
  'text/index.mjs',
  'text/observation.mjs',
  'text/README.md',
  textManifestRelativePath,
];
const requiredPackageDeclarations = [
  generatedManifestRelativePath,
  generatedContractDeclaration,
  textManifestRelativePath,
  textContractDeclaration,
];
const providerContractRoots = [
  'packages/convert-cli/schemas',
  'packages/docx-cli/contracts',
  'packages/pptx-cli/contracts',
  'packages/xlsx-cli/contracts',
  'servers/text/provider-contracts',
];
const forbiddenDependencyTerms = /(?:^|[^a-z0-9])(lucid|scenario|workitem)(?:$|[^a-z0-9])/i;
const forbiddenProviderOwnershipTerms = /(?:^|[^a-z0-9])(lucid|scenario|work[ -]?items?|review-required|waiting[ -]?review)(?:$|[^a-z0-9])/i;
const forbiddenSchemaTerms = /(?:^|[^a-z0-9])(workflow|terminal|candidate|decision)(?:s)?(?:$|[^a-z0-9])/i;
const schemaSignalKeys = new Set(['$id', 'title', 'type', 'const', 'enum', 'name', 'kind', 'status', 'operation', '$ref']);
const forbiddenRepositoryRefs = [
  /(?:^|[\\/])lucid-docs(?:$|[\\/])/i,
  /(?:^|[\\/])plugins[\\/]lucid(?:$|[\\/])/i,
  /(?:^|[\\/])scenarios(?:$|[\\/])/i,
  /(?:^|[\\/])workitems?(?:$|[\\/])/i,
  /\/Users\/[^\n"']*lucid/i,
];

const failures = [];
const notes = [];

function fail(check, message) {
  failures.push(`${check}: ${message}`);
}

function note(message) {
  notes.push(message);
}

async function exists(file) {
  try {
    await access(file, fsConstants.F_OK);
    return true;
  } catch {
    return false;
  }
}

async function readJson(file) {
  return JSON.parse(await readFile(file, 'utf8'));
}

function dependencyNamesFromPackage(packageJson) {
  return [
    packageJson.dependencies,
    packageJson.devDependencies,
    packageJson.optionalDependencies,
    packageJson.peerDependencies,
  ].flatMap(section => Object.keys(section || {}));
}

function dependencyEntriesFromPackage(packageJson) {
  return [
    packageJson.dependencies,
    packageJson.devDependencies,
    packageJson.optionalDependencies,
    packageJson.peerDependencies,
  ].flatMap(section => Object.entries(section || {}));
}

function packageNameFromLockPath(lockPath) {
  if (!lockPath.startsWith('node_modules/')) return null;
  const parts = lockPath.slice('node_modules/'.length).split('/');
  return parts[0]?.startsWith('@') ? parts.slice(0, 2).join('/') : parts[0];
}

async function checkDependencyGraph() {
  const check = 'dependency-graph';
  if (!(await exists(serverPackagePath))) {
    fail(check, `missing ${serverPackagePath}`);
    return;
  }
  if (!(await exists(serverLockPath))) {
    fail(check, `missing ${serverLockPath}`);
    return;
  }

  const packageJson = await readJson(serverPackagePath);
  const lock = await readJson(serverLockPath);
  const names = new Set(dependencyNamesFromPackage(packageJson));
  for (const [name, spec] of dependencyEntriesFromPackage(packageJson)) {
    if (forbiddenDependencyTerms.test(name)) fail(check, `forbidden package in package.json: ${name}`);
    if (forbiddenRepositoryRefs.some(pattern => pattern.test(String(spec)))) {
      fail(check, `dependency spec references a Lucid repository: ${name}=${spec}`);
    }
  }
  for (const [lockPath, entry] of Object.entries(lock.packages || {})) {
    const name = entry?.name || packageNameFromLockPath(lockPath);
    if (name) names.add(name);
    for (const field of ['resolved', 'version']) {
      if (entry?.[field] && forbiddenRepositoryRefs.some(pattern => pattern.test(String(entry[field])))) {
        fail(check, `lockfile ${field} references a Lucid repository: ${lockPath}`);
      }
    }
    for (const [dependencyName, spec] of [
      ...Object.entries(entry?.dependencies || {}),
      ...Object.entries(entry?.optionalDependencies || {}),
      ...Object.entries(entry?.peerDependencies || {}),
    ]) {
      if (forbiddenDependencyTerms.test(dependencyName)) fail(check, `forbidden transitive package in graph: ${dependencyName}`);
      if (forbiddenRepositoryRefs.some(pattern => pattern.test(String(spec)))) {
        fail(check, `lockfile dependency spec references a Lucid repository: ${dependencyName}=${spec}`);
      }
    }
  }
  for (const name of names) {
    if (forbiddenDependencyTerms.test(name)) fail(check, `forbidden package in graph: ${name}`);
  }

  const root = lock.packages?.[''];
  if (root?.name !== packageJson.name) fail(check, 'lockfile root name does not match package.json');
  if (root?.version !== packageJson.version) fail(check, 'lockfile root version does not match package.json');
  if (packageJson.private === true) fail(check, 'published Office package must not be private');
  note(`dependency graph checked: ${names.size} package names`);
}

async function checkOfficeSourceOwnership() {
  const check = 'office-source-ownership';
  const officeSource = await readFile(path.join(serverRoot, 'office', 'index.mjs'), 'utf8');
  if (forbiddenProviderOwnershipTerms.test(officeSource)) {
    fail(check, 'Office MCP source contains Lucid/scenario/work-item ownership terms');
  }
  if (/inputSchema\s*:\s*z\./.test(officeSource)
      || /const\s+(?:inputOnly|artifactInput|pathInput)\s*=\s*z\./.test(officeSource)) {
    fail(check, 'Office MCP source hand-defines an MCP input schema instead of loading a provider contract');
  }
  if (!officeSource.includes('z.fromJSONSchema(schema)')
      || !officeSource.includes('tiwater-office-provider-contract-manifest-v1.json')) {
    fail(check, 'Office MCP source does not register provider-owned input contracts from the generated manifest');
  }
  if (/\boperationType\b|\bsourceFields\b|withTempJsonFile\s*\(\s*\{\s*operations\b/.test(officeSource)
      || /runJsonCandidateChain\([^\n]*\[\s*['"]edit['"]/.test(officeSource)) {
    fail(check, 'Office MCP adapter translates a fixed provider call into a second operation request');
  }
  if (!officeSource.includes('[tool, requestPath]')) {
    fail(check, 'Office MCP adapter does not forward fixed provider calls using their published tool identity');
  }
  note('Office MCP source owns routing only; provider contracts own request shapes and fixed runtimes consume them directly');
}

async function checkTextSourceOwnership() {
  const check = 'text-source-ownership';
  const textSource = await readFile(path.join(serverRoot, 'text', 'index.mjs'), 'utf8');
  const providerSource = await readFile(path.join(serverRoot, 'text', 'observation.mjs'), 'utf8');
  if (forbiddenProviderOwnershipTerms.test(`${textSource}\n${providerSource}`)) {
    fail(check, 'Text MCP source contains Lucid/scenario/work-item ownership terms');
  }
  if (/inputSchema\s*:\s*z\./.test(textSource)) {
    fail(check, 'Text MCP source hand-defines an MCP input schema instead of loading a provider contract');
  }
  if (!textSource.includes('z.fromJSONSchema(JSON.parse(bytes.toString')
      || !textSource.includes('tiwater-text-provider-contract-manifest-v1.json')) {
    fail(check, 'Text MCP source does not register provider-owned input contracts from the generated manifest');
  }
  if (!textSource.includes("from '../_shared/large-json-result.mjs'")
      || !textSource.includes("from '../_shared/output-write-lock.mjs'")) {
    fail(check, 'Text MCP does not reuse distribution-owned artifact and output-lock helpers');
  }
  note('Text MCP owns lossless byte/line observation only and reuses the shared distribution runtime');
}

async function checkFixedRuntimeSurface() {
  const check = 'fixed-runtime-surface';
  const officeSource = await readFile(path.join(serverRoot, 'office', 'index.mjs'), 'utf8');
  const fixedNames = new Set([
    ...[...officeSource.matchAll(/\{"name":"((?:docx|xlsx)_[^"]+)"/g)].map(match => match[1]),
    ...[...officeSource.matchAll(/fixedEdit\('((?:docx|xlsx|pptx)_[^']+)'/g)].map(match => match[1]),
    ...[...officeSource.matchAll(/fixedCreate\('((?:docx|xlsx|pptx)_[^']+)'/g)].map(match => match[1]),
    ...[...officeSource.matchAll(/docxObservation\('([^']+)'/g)].map(match => match[1]),
  ]);
  const providerSources = (await Promise.all([
    'packages/docx-cli',
    'packages/xlsx-cli',
    'packages/pptx-cli',
  ].map(async relative => {
    const directory = path.join(repoRoot, relative);
    const files = (await readdir(directory)).filter(name => name.endsWith('.cs'));
    return Promise.all(files.map(name => readFile(path.join(directory, name), 'utf8')));
  }))).flat();
  for (const name of fixedNames) {
    if (!providerSources.some(source => source.includes(`"${name}"`))) {
      fail(check, `Office MCP fixed tool has no same-name provider command: ${name}`);
    }
  }

  const programSources = await Promise.all([
    'packages/docx-cli/Program.cs',
    'packages/xlsx-cli/Program.cs',
    'packages/pptx-cli/Program.cs',
  ].map(relative => readFile(path.join(repoRoot, relative), 'utf8')));
  const genericPublicCommands = /"(?:edit|list|find|read|copy-table-range|apply-format-edits|set-shape-geometry|replace-picture-image|apply-template)"\s*(?:,|=>)/;
  if (programSources.some(source => genericPublicCommands.test(source))) {
    fail(check, 'provider CLI still publishes a second generic edit/plan command');
  }
  const fixedEditSource = officeSource.slice(officeSource.indexOf('async function fixedEdit('));
  if (!fixedEditSource.includes('documentMutationFileArguments(publishedContract, args)')
      || fixedEditSource.includes("tool.startsWith('docx_')")
      || fixedEditSource.includes("tool.startsWith('xlsx_')")
      || fixedEditSource.includes("tool.startsWith('pptx_')")) {
    fail(check, 'Office MCP fixed edits do not derive current/effective paths from the published mutation contract');
  }
  note(`fixed provider runtime surface checked: ${fixedNames.size} same-name commands and no generic edit route`);
}

async function checkPackageFiles() {
  const check = 'package-files';
  const packageJson = await readJson(serverPackagePath);
  const declaredFiles = Array.isArray(packageJson.files) ? packageJson.files : [];
  if (declaredFiles.length === 0) fail(check, 'servers/package.json has no files allow-list');

  for (const declaration of requiredPackageDeclarations) {
    if (!declaredFiles.includes(declaration)) {
      fail(check, `package files allow-list must declare generated provider contracts: ${declaration}`);
    }
  }

  for (const file of requiredPackageFiles) {
    if (!(await exists(path.join(serverRoot, file)))) fail(check, `missing required source file: ${file}`);
  }

  note(`package file allow-list present: ${declaredFiles.length} declarations; npm pack is the source of truth for included files`);
}

function extractPackResult(stdout) {
  const parsed = JSON.parse(stdout);
  if (!Array.isArray(parsed) || parsed.length !== 1 || !Array.isArray(parsed[0].files)) {
    throw new Error('npm pack did not return one package manifest');
  }
  return parsed[0];
}

async function packOfficePackage(tempRoot) {
  const destination = path.join(tempRoot, 'pack');
  await mkdirIfMissing(destination);
  const { stdout } = await execFileAsync('npm', [
    'pack', '--json', '--ignore-scripts', '--pack-destination', destination,
  ], { cwd: serverRoot, maxBuffer: 4 * 1024 * 1024 });
  const manifest = extractPackResult(stdout);
  if (manifest.name !== '@tiwater/office-mcp') fail('pack-manifest', `unexpected package name: ${manifest.name}`);
  if (manifest.entryCount !== manifest.files.length) fail('pack-manifest', 'entryCount does not match files length');
  const archive = path.join(destination, manifest.filename);
  if (!(await exists(archive))) fail('pack-manifest', `pack archive missing: ${archive}`);
  return { archive, manifest };
}

async function mkdirIfMissing(directory) {
  await mkdir(directory, { recursive: true });
}

async function extractArchive(archive, destination) {
  await mkdirIfMissing(destination);
  await execFileAsync('tar', ['-xzf', archive, '-C', destination], { maxBuffer: 1024 * 1024 });
  return path.join(destination, 'package');
}

async function collectFiles(directory) {
  const output = [];
  async function visit(current) {
    for (const entry of await readdir(current, { withFileTypes: true })) {
      const absolute = path.join(current, entry.name);
      if (entry.isDirectory()) await visit(absolute);
      else output.push(absolute);
    }
  }
  await visit(directory);
  return output;
}

function isTextFile(file) {
  return /\.(?:cjs|js|json|mjs|md|schema|ts|txt|yaml|yml)$/i.test(file);
}

async function checkPackedPackage(manifest, packageRoot) {
  const check = 'pack-manifest';
  const packedPaths = new Set(manifest.files.map(file => file.path));
  for (const file of manifest.files) {
    if (path.posix.isAbsolute(file.path) || file.path.split('/').includes('..')) {
      fail(check, `pack manifest contains unsafe path: ${file.path}`);
    }
  }
  for (const required of requiredPackageFiles) {
    if (!packedPaths.has(required)) fail(check, `required file absent from pack: ${required}`);
  }
  const contractPaths = manifest.files
    .map(file => file.path)
    .filter(file => file.startsWith('office/contracts/') && file.endsWith('.schema.json'));
  const textContractPaths = manifest.files
    .map(file => file.path)
    .filter(file => file.startsWith('text/contracts/') && file.endsWith('.schema.json'));
  if (!packedPaths.has(generatedManifestRelativePath)) {
    fail(check, `generated provider contract manifest is absent from pack: ${generatedManifestRelativePath}`);
  }
  if (contractPaths.length === 0) {
    fail(check, 'Office MCP pack must contain at least one generated public contract schema');
  } else {
    note(`Office MCP pack contains ${contractPaths.length} public contract schemas`);
  }
  if (!packedPaths.has(textManifestRelativePath)) {
    fail(check, `generated Text provider contract manifest is absent from pack: ${textManifestRelativePath}`);
  }
  if (textContractPaths.length !== 2) {
    fail(check, `Text MCP pack must contain exactly two generated public contract schemas, found ${textContractPaths.length}`);
  } else {
    note('Text MCP pack contains 2 public contract schemas');
  }

  for (const file of await collectFiles(packageRoot)) {
    if (!isTextFile(file)) continue;
    const contents = await readFile(file, 'utf8');
    for (const pattern of forbiddenRepositoryRefs) {
      if (pattern.test(contents)) fail(check, `packed file references Lucid repository: ${path.relative(packageRoot, file)}`);
    }
  }
  const packedPackage = await readJson(path.join(packageRoot, 'package.json'));
  if (packedPackage.name !== '@tiwater/office-mcp') fail(check, 'packed package.json name is not @tiwater/office-mcp');
  if (packedPackage.version !== manifest.version) fail(check, 'packed package.json version does not match npm pack manifest');
  note(`pack manifest checked: ${manifest.entryCount} entries`);
}

function schemaSignals(value, location = '$', signals = []) {
  if (Array.isArray(value)) {
    value.forEach((item, index) => schemaSignals(item, `${location}[${index}]`, signals));
    return signals;
  }
  if (!value || typeof value !== 'object') return signals;
  for (const [key, child] of Object.entries(value)) {
    if (forbiddenSchemaTerms.test(key)) signals.push(`${location}.${key}`);
    if (schemaSignalKeys.has(key)) {
      const values = Array.isArray(child) ? child : [child];
      for (const item of values) {
        if (typeof item === 'string' && forbiddenSchemaTerms.test(item)) signals.push(`${location}.${key}=${item}`);
      }
    }
    schemaSignals(child, `${location}.${key}`, signals);
  }
  return signals;
}

function docxAddressContractSignals(value, location = '$', signals = []) {
  if (Array.isArray(value)) {
    value.forEach((item, index) => docxAddressContractSignals(item, `${location}[${index}]`, signals));
    return signals;
  }
  if (!value || typeof value !== 'object') return signals;
  if (value.properties?.part && value.properties?.path) {
    if (value.properties.part.pattern !== '^\\/') signals.push(`${location}.properties.part`);
    if (value.properties.path.pattern !== nativeDocxPathPattern) signals.push(`${location}.properties.path`);
  }
  for (const [key, child] of Object.entries(value)) {
    docxAddressContractSignals(child, `${location}.${key}`, signals);
  }
  return signals;
}

async function checkPublicSchemas(packageRoot) {
  const check = 'public-schemas';
  const packedFiles = (await collectFiles(packageRoot))
    .filter(file => file.endsWith('.schema.json') || (file.includes(`${path.sep}contracts${path.sep}`) && file.endsWith('.json')))
    .map(file => ({ file, root: packageRoot }));
  const sourceFiles = [];
  for (const relativeRoot of providerContractRoots) {
    const root = path.join(repoRoot, relativeRoot);
    for (const file of await collectFiles(root).catch(() => [])) {
      if (file.endsWith('.schema.json') || (file.includes(`${path.sep}contracts${path.sep}`) && file.endsWith('.json'))) {
        sourceFiles.push({ file, root: repoRoot });
      }
    }
  }
  const contractFiles = [...sourceFiles, ...packedFiles];
  if (contractFiles.length === 0) {
    note('no packaged public schema/contract files to inspect');
    return;
  }
  for (const { file, root } of contractFiles) {
    let schema;
    try {
      schema = JSON.parse(await readFile(file, 'utf8'));
    } catch (error) {
      fail(check, `invalid JSON schema ${path.relative(root, file)}: ${error.message}`);
      continue;
    }
    const signals = schemaSignals(schema);
    for (const signal of signals) fail(check, `${path.relative(root, file)} contains business type signal ${signal}`);
    if (path.basename(file).startsWith('docx_')) {
      for (const signal of docxAddressContractSignals(schema)) {
        fail(check, `${path.relative(root, file)} has an incomplete native DOCX address contract at ${signal}`);
      }
    }
    const text = JSON.stringify(schema);
    if (forbiddenSchemaTerms.test(text)) fail(check, `${path.relative(root, file)} hits schema keyword backstop`);
  }
  note(`public schemas checked: ${contractFiles.length} (${sourceFiles.length} provider, ${packedFiles.length} packed)`);
}

async function sha256(file) {
  return createHash('sha256').update(await readFile(file)).digest('hex');
}

function exactKeys(value, keys) {
  return value && typeof value === 'object' && !Array.isArray(value)
    && Object.keys(value).sort().join('\0') === [...keys].sort().join('\0');
}

function safeRelativePath(value) {
  return typeof value === 'string'
    && value.length > 0
    && value === path.posix.normalize(value)
    && !path.posix.isAbsolute(value)
    && !value.split('/').includes('..');
}

function providerSourcePath(source) {
  if (!safeRelativePath(source)) return null;
  const absolute = path.resolve(repoRoot, ...source.split('/'));
  const allowed = providerContractRoots.some(root => {
    const rootPath = path.resolve(repoRoot, root);
    const relative = path.relative(rootPath, absolute);
    return relative === '' || (!relative.startsWith(`..${path.sep}`) && relative !== '..' && !path.isAbsolute(relative));
  });
  return allowed ? absolute : null;
}

async function checkGeneratedManifest(packageRoot, toolNames, packageManifest) {
  const check = 'generated-manifest';
  const manifestPath = path.join(packageRoot, generatedManifestRelativePath);
  if (!(await exists(manifestPath))) {
    fail(check, `generated provider contract manifest is missing: ${generatedManifestRelativePath}`);
    return;
  }

  let manifest;
  try {
    manifest = await readJson(manifestPath);
  } catch (error) {
    fail(check, `generated provider contract manifest is invalid JSON: ${error.message}`);
    return;
  }
  if (!exactKeys(manifest, ['schema', 'provider', 'tools'])) {
    fail(check, 'generated provider contract manifest must contain exactly schema, provider, and tools');
    return;
  }
  if (manifest.schema !== 'tiwater.office-provider-contract-manifest/v1') {
    fail(check, `unexpected generated manifest schema: ${manifest.schema}`);
  }
  for (const signal of schemaSignals(manifest)) {
    fail(check, `generated manifest contains business type signal ${signal}`);
  }
  if (forbiddenSchemaTerms.test(JSON.stringify(manifest))) {
    fail(check, 'generated manifest hits schema keyword backstop');
  }
  if (!exactKeys(manifest.provider, ['id', 'version'])
      || manifest.provider.id !== packageManifest.name
      || manifest.provider.version !== packageManifest.version) {
    fail(check, 'generated manifest provider identity does not match the packed package');
  }
  if (!Array.isArray(manifest.tools) || manifest.tools.length === 0) {
    fail(check, 'generated provider contract manifest must declare at least one MCP tool');
    return;
  }

  const declaredNames = manifest.tools.map(entry => entry?.name);
  if (declaredNames.some(name => typeof name !== 'string' || name.length === 0)
      || new Set(declaredNames).size !== declaredNames.length) {
    fail(check, 'generated manifest tool names must be non-empty and unique');
  }
  const expectedNames = [...new Set(toolNames.map(tool => tool?.name))].sort();
  const actualNames = [...declaredNames].sort();
  if (expectedNames.join('\0') !== actualNames.join('\0')) {
    fail(check, `generated manifest tools do not exactly match MCP tools/list (${actualNames.join(', ')})`);
  }

  const packedContractFiles = new Set((await collectFiles(packageRoot))
    .filter(file => file.endsWith('.schema.json')
      && file.includes(`${path.sep}office${path.sep}contracts${path.sep}`))
    .map(file => path.relative(packageRoot, file).split(path.sep).join('/')));
  const referencedContractFiles = new Set();
  let fileRoleCount = 0;
  for (const entry of manifest.tools) {
    if (!exactKeys(entry, ['name', 'providerContract', 'inputContract'])
        || !exactKeys(entry.providerContract, ['source', 'sha256'])
        || !exactKeys(entry.inputContract, ['path', 'sha256'])) {
      fail(check, `tool ${entry?.name || '(unnamed)'} must declare exactly providerContract.source/sha256 and inputContract.path/sha256`);
      continue;
    }
    const { source, sha256: declaredHash } = entry.providerContract;
    const { path: packagedPath, sha256: packagedHash } = entry.inputContract;
    const sourcePath = providerSourcePath(source);
    if (!sourcePath) {
      fail(check, `tool ${entry.name} provider contract source is not an allowed provider schema: ${source}`);
      continue;
    }
    if (!source.includes('/mcp-input/') || path.posix.basename(source) !== `${entry.name}.schema.json`) {
      fail(check, `tool ${entry.name} provider source is not its bound MCP input contract: ${source}`);
      continue;
    }
    if (!(await exists(sourcePath))) {
      fail(check, `tool ${entry.name} provider contract source is missing: ${source}`);
      continue;
    }
    if (!/^[a-f0-9]{64}$/.test(declaredHash) || !/^[a-f0-9]{64}$/.test(packagedHash)) {
      fail(check, `tool ${entry.name} provider/input contract hashes must be SHA-256 values`);
      continue;
    }
    if (!safeRelativePath(packagedPath) || !packagedPath.startsWith('office/contracts/') || !packagedPath.endsWith('.schema.json')) {
      fail(check, `tool ${entry.name} provider contract package path is invalid: ${packagedPath}`);
      continue;
    }
    const packedContractPath = path.join(packageRoot, ...packagedPath.split('/'));
    if (!(await exists(packedContractPath))) {
      fail(check, `tool ${entry.name} provider contract is absent from pack: ${packagedPath}`);
      continue;
    }
    const [sourceHash, actualPackedHash] = await Promise.all([sha256(sourcePath), sha256(packedContractPath)]);
    if (declaredHash !== sourceHash) fail(check, `tool ${entry.name} hash does not match source schema: ${source}`);
    if (packagedHash !== actualPackedHash) fail(check, `tool ${entry.name} input contract hash does not match packaged contract: ${packagedPath}`);
    if (declaredHash !== packagedHash) fail(check, `tool ${entry.name} provider and packaged input contract hashes differ`);
    let packagedInputSchema;
    try {
      packagedInputSchema = await readJson(packedContractPath);
    } catch (error) {
      fail(check, `tool ${entry.name} packaged input contract is not valid JSON: ${error.message}`);
      continue;
    }
    if (packagedInputSchema.type !== 'object') {
      fail(check, `tool ${entry.name} MCP input contract root must be an object`);
    }
    fileRoleCount += checkFileArgumentRoles(packagedInputSchema, entry.name, check);
    const expectedInputSchema = canonicalMcpInputSchema(toolNames.find(tool => tool?.name === entry.name)?.inputSchema);
    const actualInputSchema = canonicalMcpInputSchema(packagedInputSchema);
    if (expectedInputSchema === null) {
      fail(check, `tool ${entry.name} MCP tools/list entry has no inputSchema object`);
    } else if (JSON.stringify(expectedInputSchema) !== JSON.stringify(actualInputSchema)) {
      fail(check, `tool ${entry.name} inputSchema does not canonically match packaged input contract: ${packagedPath}`);
    }
    referencedContractFiles.add(packagedPath);
  }
  for (const packagedPath of packedContractFiles) {
    if (!referencedContractFiles.has(packagedPath)) {
      fail(check, `packaged public contract is not declared by any MCP tool: ${packagedPath}`);
    }
  }
  note(`generated manifest checked: ${manifest.tools.length} MCP tools, ${referencedContractFiles.size} provider contracts, and ${fileRoleCount} file arguments`);
}

function checkFileArgumentRoles(schema, toolName, check) {
  let count = 0;
  function visit(node, location, propertyName = '', insideComposition = false) {
    if (!node || typeof node !== 'object' || Array.isArray(node)) return;
    const declaredRole = node[fileRoleKey];
    const declaredEffect = node[fileEffectKey];
    if (declaredRole !== undefined) {
      if (insideComposition) {
        fail(check, `tool ${toolName} hides ${fileRoleKey} inside a schema composition at ${location}`);
      } else if (node.type !== 'string' || !['read', 'write'].includes(declaredRole)) {
        fail(check, `tool ${toolName} has invalid ${fileRoleKey} at ${location}`);
      } else {
        count += 1;
      }
    }
    if (declaredEffect !== undefined
      && (declaredRole !== 'write' || typeof declaredEffect !== 'boolean')) {
      fail(check, `tool ${toolName} has invalid ${fileEffectKey} at ${location}`);
    }
    const expectedRole = node.type === 'string' ? expectedFileRolesByProperty.get(propertyName) : undefined;
    if (expectedRole && declaredRole !== expectedRole) {
      fail(check, `tool ${toolName} must declare ${location} as a ${expectedRole} file argument`);
    }
    const expectedEffect = expectedFileEffectsByProperty.get(propertyName);
    if (expectedEffect !== undefined && declaredEffect !== expectedEffect) {
      fail(check, `tool ${toolName} must declare ${location} ${fileEffectKey} as ${expectedEffect}`);
    }
    for (const [name, child] of Object.entries(node.properties || {})) {
      visit(child, `${location}.properties.${name}`, name, insideComposition);
    }
    if (node.items) visit(node.items, `${location}.items`, propertyName, insideComposition);
    for (const keyword of ['allOf', 'anyOf', 'oneOf']) {
      for (const [index, child] of (node[keyword] || []).entries()) {
        visit(child, `${location}.${keyword}[${index}]`, propertyName, true);
      }
    }
  }
  visit(schema, '$');
  if (count === 0) fail(check, `tool ${toolName} declares no file arguments`);
  return count;
}

const allowedMcpSchemaMetadata = new Map([
  ['$schema', new Set([
    'http://json-schema.org/draft-07/schema#',
    'https://json-schema.org/draft/2020-12/schema',
  ])],
]);

function canonicalMcpInputSchema(value, location = '$') {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  return canonicalMcpValue(value, location);
}

function canonicalMcpValue(value, location) {
  if (Array.isArray(value)) return value.map((item, index) => canonicalMcpValue(item, `${location}[${index}]`));
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(Object.entries(value)
    .filter(([key, child]) => {
      if (location !== '$' || !allowedMcpSchemaMetadata.has(key)) return true;
      return !allowedMcpSchemaMetadata.get(key).has(child);
    })
    .sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)
    .map(([key, child]) => {
      if (key === 'required' && Array.isArray(child) && child.every(item => typeof item === 'string')) {
        return [key, [...child].sort()];
      }
      return [key, canonicalMcpValue(child, `${location}.${key}`)];
    }));
}

function readLines(buffer) {
  return buffer.toString('utf8').split(/\r?\n/).filter(Boolean);
}

async function initializeInstalledMcp(executable, installRoot) {
  return new Promise((resolve, reject) => {
    const child = spawn(executable, [], {
      cwd: installRoot,
      env: { ...process.env, PATH: `${path.join(installRoot, 'node_modules', '.bin')}${path.delimiter}${process.env.PATH || ''}` },
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    let settled = false;
    let toolsListRequested = false;
    let serverInstructions = '';
    const finish = (result, error = null) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      child.kill('SIGTERM');
      if (error) reject(error);
      else resolve({ ...result, stderr });
    };
    const timer = setTimeout(() => finish(null, new Error('MCP initialize/tools-list timed out after 12 seconds')), 12_000);
    child.stdout.on('data', chunk => {
      stdout += chunk.toString('utf8');
      for (const line of readLines(stdout)) {
        let message;
        try { message = JSON.parse(line); } catch { continue; }
        if (message?.id === 1 && !toolsListRequested) {
          if (message.error) {
            finish(null, new Error(`MCP initialize returned error: ${JSON.stringify(message.error)}`));
          } else if (message.result && typeof message.result === 'object') {
            serverInstructions = typeof message.result.instructions === 'string'
              ? message.result.instructions : '';
            toolsListRequested = true;
            child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }) + '\n');
          }
        } else if (message?.id === 2) {
          if (message.error) {
            finish(null, new Error(`MCP tools/list returned error: ${JSON.stringify(message.error)}`));
          } else if (Array.isArray(message.result?.tools)) {
            const tools = message.result.tools;
            if (tools.length === 0) {
              finish(null, new Error('MCP tools/list returned no tools'));
            } else if (tools.some(tool => !tool || typeof tool.name !== 'string' || !canonicalMcpInputSchema(tool.inputSchema))) {
              finish(null, new Error('MCP tools/list returned a tool without a complete inputSchema'));
            } else if (new Set(tools.map(tool => tool.name)).size !== tools.length) {
              finish(null, new Error('MCP tools/list returned duplicate tool names'));
            } else {
              finish({ initialized: true, tools, serverInstructions });
            }
          } else {
            finish(null, new Error('MCP tools/list returned no tools array'));
          }
        }
      }
      if (stdout.length > 2 * 1024 * 1024) finish(null, new Error('MCP initialize emitted excessive stdout'));
    });
    child.stderr.on('data', chunk => { stderr += chunk.toString('utf8'); });
    child.on('error', error => finish(null, error));
    child.on('exit', (code, signal) => {
      if (!settled) finish(null, new Error(`MCP process exited before initialize response (code=${code}, signal=${signal})`));
    });
    child.stdin.write(JSON.stringify({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2025-06-18',
        capabilities: {},
        clientInfo: { name: 'release-boundary-gate', version: '1.0.0' },
      },
    }) + '\n');
  });
}

async function smokeInstalledPackage(archive, tempRoot) {
  const check = 'isolated-smoke';
  const installRoot = path.join(tempRoot, 'unrelated-install');
  await mkdirIfMissing(installRoot);
  if (!path.relative(repoRoot, installRoot).startsWith('..')) {
    fail(check, `smoke directory is inside the Lucid/provider repository: ${installRoot}`);
    return { officeTools: [], textTools: [] };
  }
  await execFileAsync('npm', [
    'install', '--ignore-scripts', '--no-audit', '--no-fund', '--package-lock=false', '--prefix', installRoot, archive,
  ], { cwd: installRoot, maxBuffer: 8 * 1024 * 1024 });

  const executable = path.join(installRoot, 'node_modules', '.bin', 'tiwater-office-mcp');
  if (!(await exists(executable))) {
    fail(check, 'installed package did not expose tiwater-office-mcp executable');
    return { officeTools: [], textTools: [] };
  }

  const response = await initializeInstalledMcp(executable, installRoot);
  if (!response.initialized) fail(check, 'MCP initialize did not complete');
  if (!response.serverInstructions.includes('A read-only output path is an immutable artifact identity')
      || !response.serverInstructions.includes('an identical request may replay it')
      || !response.serverInstructions.includes('every different request uses a different path')) {
    fail(check, 'MCP instructions do not publish immutable read artifact path semantics');
  }
  if (!response.serverInstructions.includes('Every mutation receiptOutput is a new immutable receipt identity')
      || !response.serverInstructions.includes('when the same document object is updated again')) {
    fail(check, 'MCP instructions do not publish per-call mutation receipt identity');
  }
  if (!response.serverInstructions.includes('Every native DOCX address belongs only to the exact input file')
      || !response.serverInstructions.includes('must not be reused with another DOCX')) {
    fail(check, 'MCP instructions do not publish native DOCX address scope');
  }
  const textExecutable = path.join(installRoot, 'node_modules', '.bin', 'tiwater-text-mcp');
  if (!(await exists(textExecutable))) {
    fail(check, 'installed package did not expose tiwater-text-mcp executable');
    return { officeTools: response.tools || [], textTools: [] };
  }
  const textResponse = await initializeInstalledMcp(textExecutable, installRoot);
  if (!textResponse.serverInstructions.includes('exact supported plain-text bytes')
      || !textResponse.serverInstructions.includes('Callers own all interpretation and business meaning')) {
    fail(check, 'Text MCP instructions do not publish technical-only plain-text observation ownership');
  }
  note(`isolated Office and Text MCP initialize/tools-list completed (${response.tools?.length || 0} Office, ${textResponse.tools?.length || 0} Text)${response.stderr.trim() || textResponse.stderr.trim() ? ' with stderr output' : ''}`);
  return { officeTools: response.tools || [], textTools: textResponse.tools || [] };
}

async function checkTextPublishedSurface(packageRoot, tools, packageManifest) {
  const check = 'text-published-surface';
  const manifestPath = path.join(packageRoot, textManifestRelativePath);
  const manifest = await readJson(manifestPath);
  if (!exactKeys(manifest, ['schema', 'provider', 'tools'])
      || manifest.schema !== 'tiwater.text-provider-contract-manifest/v1') {
    fail(check, 'Text manifest has an unexpected envelope or schema identity');
    return;
  }
  if (!exactKeys(manifest.provider, ['id', 'version'])
      || manifest.provider.id !== packageManifest.name
      || manifest.provider.version !== packageManifest.version) {
    fail(check, 'Text manifest provider identity does not match the packed distribution');
  }
  const expectedNames = ['text_inspect', 'text_read_lines'];
  const manifestNames = manifest.tools.map(entry => entry?.name).sort();
  const toolNames = tools.map(tool => tool?.name).sort();
  if (manifestNames.join('\0') !== expectedNames.join('\0')
      || toolNames.join('\0') !== expectedNames.join('\0')) {
    fail(check, 'Text manifest and MCP tools/list must expose exactly text_inspect and text_read_lines');
  }
  for (const entry of manifest.tools) {
    if (!exactKeys(entry, ['name', 'providerContract', 'inputContract'])
        || !exactKeys(entry.providerContract, ['source', 'sha256'])
        || !exactKeys(entry.inputContract, ['path', 'sha256'])) {
      fail(check, `Text tool ${entry?.name || '(unnamed)'} has an incomplete hash binding`);
      continue;
    }
    const sourcePath = path.join(repoRoot, ...entry.providerContract.source.split('/'));
    const packedPath = path.join(packageRoot, ...entry.inputContract.path.split('/'));
    if (!entry.providerContract.source.startsWith('servers/text/provider-contracts/')
        || entry.inputContract.path !== `text/contracts/${entry.name}.schema.json`
        || !(await exists(sourcePath))
        || !(await exists(packedPath))) {
      fail(check, `Text tool ${entry.name} is not bound to its provider-owned and packed schema`);
      continue;
    }
    const [sourceHash, packedHash] = await Promise.all([sha256(sourcePath), sha256(packedPath)]);
    if (sourceHash !== entry.providerContract.sha256
        || packedHash !== entry.inputContract.sha256
        || sourceHash !== packedHash) {
      fail(check, `Text tool ${entry.name} contract hashes do not bind identical source and packed bytes`);
    }
  }
  for (const tool of tools) {
    const annotations = tool.annotations || {};
    if (annotations.readOnlyHint !== true || annotations.idempotentHint !== true
        || annotations.destructiveHint !== false || annotations.openWorldHint !== false) {
      fail(check, `${tool.name} does not publish closed-world read-only annotations`);
    }
  }
  const inspect = tools.find(tool => tool.name === 'text_inspect');
  const read = tools.find(tool => tool.name === 'text_read_lines');
  const inspectRequired = inspect?.inputSchema?.required || [];
  const readRequired = read?.inputSchema?.required || [];
  if (!['input', 'returnContent', 'output'].every(name => inspectRequired.includes(name))
      || inspect?.inputSchema?.properties?.input?.[fileRoleKey] !== 'read'
      || inspect?.inputSchema?.properties?.output?.[fileRoleKey] !== 'write') {
    fail(check, 'text_inspect does not require its explicit input, content channel, and durable artifact');
  }
  if (!['input', 'offset', 'returnContent'].every(name => readRequired.includes(name))
      || read?.inputSchema?.properties?.offset?.minimum !== 0
      || Object.hasOwn(read?.inputSchema?.properties ?? {}, 'limit')
      || !(read?.description || '').includes('provider chooses the bounded page size')) {
    fail(check, 'text_read_lines does not publish provider-sized zero-based paging');
  }
  const inspectOutput = inspect?.outputSchema;
  const readOutput = read?.outputSchema;
  if (inspectOutput?.properties?.identity?.properties?.openingLines?.maxItems !== 8
      || !inspectOutput?.properties?.identity?.required?.includes('lineCount')
      || readOutput?.properties?.content?.properties?.lines?.maxItems !== 200
      || !['remaining', 'nextOffset'].every(name => readOutput?.properties?.summary?.required?.includes(name))) {
    fail(check, 'Text outputs do not publish bounded opening identity and explicit continuation facts');
  }
  note('Text manifest, schemas, annotations, bounded outputs, and MCP surface are hash-bound and orthogonal');
}

function checkEvidenceRoleMetadata(officeTools, textTools) {
  const check = 'provider-evidence-role-metadata';
  const expected = new Map([
    ['docx_inspect', 'document-observation'],
    ['xlsx_inspect', 'document-observation'],
    ['pptx_inspect', 'document-observation'],
    ['text_inspect', 'document-observation'],
    ['docx_export_json', 'final-readback'],
    ['xlsx_export_json', 'final-readback'],
    ['pptx_export_json', 'final-readback'],
    ['office_render_pdf', 'native-render'],
  ]);
  const tools = [...officeTools, ...textTools];
  for (const tool of tools) {
    const expectedRole = expected.get(tool.name);
    const hasMetadata = tool?._meta?.[evidenceRoleMetadataKey] !== undefined;
    if (!expectedRole) {
      if (hasMetadata) fail(check, `${tool.name} publishes an unexpected evidence role`);
      continue;
    }
    try {
      assertEvidenceToolContract(tool, expectedRole);
    } catch (error) {
      fail(check, error.message);
    }
    expected.delete(tool.name);
  }
  if (expected.size > 0) fail(check, `missing role-bearing tools: ${[...expected.keys()].join(', ')}`);
  note('Office and Text publish exact versioned evidence roles derived from annotations, file bindings, and output schemas');
}

function checkEffectKindMetadata(officeTools, textTools) {
  const check = 'provider-effect-kind-metadata';
  const expectedSpecialKinds = new Map([
    ['docx_create', 'document-create'],
    ['office_render_pdf', 'native-render'],
    ['xlsx_convert_legacy', 'source-conversion'],
  ]);
  let effectfulCount = 0;
  for (const tool of officeTools) {
    const expectedKind = expectedSpecialKinds.get(tool.name);
    try {
      const kind = assertEffectKindToolContract(tool, expectedKind);
      if (kind) effectfulCount += 1;
      if (kind && !expectedKind && kind !== 'document-mutation') {
        fail(check, `${tool.name} publishes unexpected effect kind ${kind}`);
      }
    } catch (error) {
      fail(check, error.message);
    }
    expectedSpecialKinds.delete(tool.name);
  }
  if (expectedSpecialKinds.size > 0) {
    fail(check, `missing special effect tools: ${[...expectedSpecialKinds.keys()].join(', ')}`);
  }
  for (const tool of textTools) {
    if (tool?._meta?.[effectKindMetadataKey] !== undefined) {
      fail(check, `${tool.name} publishes an Office effect kind`);
    }
  }
  if (effectfulCount === 0) fail(check, 'Office surface publishes no effect-bearing tools');

  const mutation = officeTools.find(tool =>
    tool?._meta?.[effectKindMetadataKey]?.kind === 'document-mutation');
  const read = officeTools.find(tool => tool?.annotations?.readOnlyHint === true);
  const render = officeTools.find(tool =>
    tool?._meta?.[effectKindMetadataKey]?.kind === 'native-render');
  const invalid = [
    mutation && { ...mutation, _meta: Object.fromEntries(Object.entries(mutation._meta)
      .filter(([key]) => key !== effectKindMetadataKey)) },
    read && { ...read, _meta: {
      ...(read._meta || {}),
      [effectKindMetadataKey]: {
        schema: 'tiwater.provider-effect-kind/v1', kind: 'document-mutation',
      },
    } },
    render && { ...render, _meta: {
      ...render._meta,
      [effectKindMetadataKey]: {
        schema: 'tiwater.provider-effect-kind/v1', kind: 'document-mutation',
      },
    } },
    mutation && { ...mutation, inputSchema: {
      ...mutation.inputSchema,
      properties: Object.fromEntries(Object.entries(mutation.inputSchema.properties || {})
        .map(([name, schema]) => [name, {
          ...schema,
          'x-tiwater-document-revision-role': undefined,
        }])),
    } },
  ].filter(Boolean);
  for (const tool of invalid) {
    let rejected = false;
    try { assertEffectKindToolContract(tool); } catch { rejected = true; }
    if (!rejected) fail(check, `known-bad effect metadata did not fail closed for ${tool.name}`);
  }
  note(`${effectfulCount} Office effect-bearing tools publish one exact versioned kind; Text remains read-only and orthogonal`);
}

function checkSourceBoundObservationOutputs(tools) {
  const check = 'source-bound-observation-output';
  const inspectSchema = tools.find(tool => tool?.name === 'docx_inspect')?.outputSchema;
  const inspectSource = inspectSchema?.properties?.source;
  if (!inspectSchema?.required?.includes('source')
      || inspectSource?.type !== 'object'
      || !['path', 'sha256', 'bytes'].every(key => inspectSource.required?.includes(key))) {
    fail(check, 'docx_inspect does not publish an exact source artifact identity');
  }
  const largeResultNames = [
    'docx_compare', 'docx_export_json', 'docx_validate', 'docx_validate_font_policy',
    'docx_validate_toc_style_policy', 'xlsx_inspect', 'xlsx_export_json', 'xlsx_read_range', 'xlsx_validate',
    'pptx_inspect', 'pptx_export_json', 'pptx_read_slide', 'pptx_read_shape', 'pptx_validate',
  ];
  for (const name of largeResultNames) {
    const sources = tools.find(tool => tool?.name === name)?.outputSchema?.properties?.sources;
    const item = sources?.items;
    if (sources?.type !== 'array'
        || sources.minItems !== 1
        || sources.maxItems !== 2
        || item?.type !== 'object'
        || !['path', 'sha256', 'bytes'].every(key => item.required?.includes(key))) {
      fail(check, `${name} does not publish exact source artifact identities`);
    }
  }
  note(`${largeResultNames.length + 1} large-result tools bind their exact source artifacts`);
}

function checkLargeResultChannels(tools) {
  const check = 'large-result-channels';
  const names = [
    'docx_compare', 'docx_export_json', 'docx_validate', 'docx_validate_font_policy',
    'docx_validate_toc_style_policy', 'xlsx_inspect', 'xlsx_export_json', 'xlsx_read_range', 'xlsx_validate',
    'pptx_inspect', 'pptx_export_json', 'pptx_read_slide', 'pptx_read_shape', 'pptx_validate',
  ];
  for (const name of names) {
    const tool = tools.find(entry => entry?.name === name);
    const input = tool?.inputSchema;
    const output = tool?.outputSchema;
    if (input?.properties?.returnContent?.type !== 'boolean'
        || !input.required?.includes('returnContent')
        || input.required?.includes('output')
        || input.properties?.output?.type !== 'string') {
      fail(check, `${name} does not publish independent returnContent and output inputs`);
    }
    if (!output?.required?.includes('returnContent')
        || !output.required?.includes('artifact')
        || !output.required?.includes('receipt')
        || output.properties?.returnContent?.type !== 'boolean'
        || output.properties?.artifact?.anyOf?.length !== 2
        || output.properties?.receipt?.type !== 'object') {
      fail(check, `${name} does not publish the common large-result receipt`);
    }
    const description = tool?.description || '';
    if (!description.includes('Set returnContent true')
        || !description.includes('Provide output')
        || !description.includes('independent')
        || !description.includes('at least one is required')) {
      fail(check, `${name} does not explain the two independent result choices`);
    }
  }
  note(`${names.length} large-result tools share one bounded-return and file-output contract`);
}

function checkDocxMergedCellDescriptions(tools) {
  const check = 'docx-merged-cell-descriptions';
  const readDescription = tools.find(tool => tool?.name === 'docx_read_table')?.description || '';
  const narrowRead = tools.find(tool => tool?.name === 'docx_read_object');
  const narrowDescription = narrowRead?.description || '';
  const narrowObject = narrowRead?.outputSchema?.$defs?.__schema0?.properties?.object;
  const setDescription = tools.find(tool => tool?.name === 'docx_set_text')?.description || '';
  const mergeDescription = tools.find(tool => tool?.name === 'docx_merge_cells')?.description || '';
  const setTable = tools.find(tool => tool?.name === 'docx_set_table');
  const setTableDescription = setTable?.description || '';
  const setTableCell = setTable?.inputSchema?.properties?.rows?.items?.properties?.cells?.items;
  if (!readDescription.includes('vertical-merge restart owns one logical value')
      || !readDescription.includes('continue cell points to verticalMergeOwner')
      || !readDescription.includes('does not repeat that value inline')) {
    fail(check, 'docx_read_table does not explain vertical-merge logical-cell identity');
  }
  if (!narrowDescription.includes('vertical-merge owner')
      || !narrowDescription.includes('resolving the restart cell value')
      || narrowObject?.properties?.verticalMergeOwner?.type !== 'object'
      || narrowObject?.properties?.logicalText?.type !== 'string') {
    fail(check, 'docx_read_object does not expose narrow merged-cell logical identity');
  }
  if (!setDescription.includes('restart cell rather than a continue cell')
      || !setDescription.includes('does not insert objects, change table structure')) {
    fail(check, 'docx_set_text does not explain merged-cell and structural non-goals');
  }
  if (!mergeDescription.includes('one-column, multi-row rectangle creates a vertical merge')
      || !mergeDescription.includes('All selected cell content moves into the top-left owner')) {
    fail(check, 'docx_merge_cells does not explain vertical grouping and content ownership');
  }
  if (!setTableDescription.includes('Each explicit cell occupies contiguous columns')
      || !setTableDescription.includes('may span logical rows')
      || !setTableDescription.includes('exact native sourceSelections')
      || !setTableDescription.includes('exposes no intermediate document')
      || !setTableDescription.includes('does not select source rows')
      || setTableCell?.properties?.rowSpan?.type !== 'integer'
      || setTableCell?.properties?.rowSpan?.minimum !== 1
      || !setTableCell?.required?.includes('text')
      || setTableCell?.properties?.sourceInput?.['x-tiwater-file-role'] !== 'read'
      || setTableCell?.properties?.sourceSelections?.type !== 'array'
      || Object.hasOwn(setTableCell?.properties ?? {}, 'verticalMerge')) {
    fail(check, 'docx_set_table does not expose one atomic explicit table input');
  }
  note('DOCX table reads expose native merges while docx_set_table accepts one explicit atomic result');
}

function checkXlsxRangeReadContract(tools) {
  const check = 'xlsx-range-read-contract';
  const tool = tools.find(entry => entry?.name === 'xlsx_read_range');
  const input = tool?.inputSchema;
  const output = tool?.outputSchema;
  const page = output?.properties?.content;
  const summary = output?.properties?.summary;
  const description = tool?.description || '';
  if (!['input', 'sheet', 'range', 'returnContent'].every(name => input?.required?.includes(name))
      || input?.properties?.range?.type !== 'string'
      || input?.properties?.offset?.minimum !== 0
      || Object.hasOwn(input?.properties ?? {}, 'limit')) {
    fail(check, 'xlsx_read_range does not require one explicit native range with provider-owned page size');
  }
  if (page?.properties?.cells?.type !== 'array'
      || page?.properties?.cells?.maxItems !== 256
      || page?.properties?.cells?.items?.properties?.physical?.type !== 'boolean'
      || page?.properties?.cells?.items?.properties?.mergeOwner?.anyOf?.length !== 2
      || summary?.properties?.remaining?.type !== 'integer'
      || summary?.properties?.nextOffset?.anyOf?.length !== 2) {
    fail(check, 'xlsx_read_range output does not expose bounded native cell facts and continuation');
  }
  if (!description.includes('row-major cell offset')
      || !description.includes('provider chooses the bounded page size')
      || !description.includes('largest leading cell page that fits the response limit')
      || !description.includes('physical presence')
      || !description.includes('remaining cells and the next offset')
      || !description.includes('does not infer regions, headers, records, field meanings, or business mappings')) {
    fail(check, 'xlsx_read_range does not publish its native paging semantics and semantic non-goals');
  }
  note('XLSX range reads expose one bounded native cell page without business inference');
}

function checkDocxCreateContract(tools) {
  const check = 'docx-create-contract';
  const tool = tools.find(entry => entry?.name === 'docx_create');
  const input = tool?.inputSchema;
  const output = tool?.outputSchema;
  const properties = input?.properties ?? {};
  if (!['output', 'receiptOutput'].every(name => input?.required?.includes(name))
      || Object.hasOwn(properties, 'input')
      || properties.output?.[fileRoleKey] !== 'write'
      || properties.output?.[fileEffectKey] === false
      || properties.receiptOutput?.[fileRoleKey] !== 'write'
      || properties.receiptOutput?.[fileEffectKey] !== false) {
    fail(check, 'docx_create does not publish one new document and one non-effect receipt');
  }
  if (output?.properties?.output?.type !== 'object'
      || output?.properties?.receipt?.type !== 'object'
      || output?.properties?.summary?.properties?.pass?.const !== true
      || tool?._meta?.[effectKindMetadataKey]?.kind !== 'document-create') {
    fail(check, 'docx_create does not publish exact creation result and effect metadata');
  }
  const description = tool?.description || '';
  if (!description.includes('minimal standards-valid DOCX')
      || !description.includes('populated incrementally with the ordinary DOCX object operations')
      || !description.includes('chooses no business wording, template, layout mapping, or target structure')) {
    fail(check, 'docx_create does not publish its composition role and semantic non-goals');
  }
  note('DOCX creation supplies only a minimal current document for ordinary incremental object operations');
}

function checkPptxBoundedReadContracts(tools) {
  const check = 'pptx-bounded-read-contracts';
  const slide = tools.find(entry => entry?.name === 'pptx_read_slide');
  const shape = tools.find(entry => entry?.name === 'pptx_read_shape');
  const slideInput = slide?.inputSchema;
  const shapeInput = shape?.inputSchema;
  const slideShapes = slide?.outputSchema?.properties?.content?.properties?.slide?.properties?.shapes;
  const shapeSegments = shape?.outputSchema?.properties?.content?.properties?.segments;
  if (!['input', 'slideNumber', 'returnContent'].every(name => slideInput?.required?.includes(name))
      || slideInput?.properties?.slideNumber?.minimum !== 1
      || slideInput?.properties?.offset?.minimum !== 0
      || Object.hasOwn(slideInput?.properties ?? {}, 'limit')
      || slideShapes?.maxItems !== 8
      || slideShapes?.items?.properties?.textPreview?.maxLength !== 240
      || slideShapes?.items?.properties?.textLength?.type !== 'integer') {
    fail(check, 'pptx_read_slide does not expose one compact bounded native shape index');
  }
  if (!['input', 'slideNumber', 'shapeId', 'returnContent'].every(name => shapeInput?.required?.includes(name))
      || shapeInput?.properties?.slideNumber?.minimum !== 1
      || shapeInput?.properties?.shapeId?.minimum !== 1
      || shapeInput?.properties?.offset?.minimum !== 0
      || Object.hasOwn(shapeInput?.properties ?? {}, 'limit')
      || shapeSegments?.maxItems !== 4
      || shapeSegments?.items?.properties?.text?.maxLength !== 160
      || shapeSegments?.items?.properties?.runIndex?.type !== 'integer'
      || shapeSegments?.items?.properties?.textOffset?.type !== 'integer') {
    fail(check, 'pptx_read_shape does not expose bounded native text and formatting segments');
  }
  if (!(slide?.description || '').includes('provider chooses the bounded page size')
      || !(shape?.description || '').includes('provider chooses the bounded page size')
      || !(slide?.description || '').includes('does not select templates, assign business roles, infer repairs')
      || !(shape?.description || '').includes('does not choose formatting, derive repairs')) {
    fail(check, 'PPTX bounded reads do not publish their semantic non-goals');
  }
  note('PPTX slide and shape reads expose compact native paging without business inference');
}

function checkDocxTableStreamingContract(tools) {
  const check = 'docx-table-streaming-contract';
  const tool = tools.find(entry => entry?.name === 'docx_read_table');
  const description = tool?.description || '';
  const input = tool?.inputSchema;
  const row = tool?.outputSchema?.properties?.rows?.items;
  const cell = row?.properties?.cells?.items;
  const receipt = tool?.outputSchema?.properties?.receipt;
  if (!description.includes('retain every remaining row')
      || !description.includes('largest compact inline page')
      || !description.includes('passing receipt.nextContinuation unchanged')
      || !description.includes('cannot be predicted or read in parallel')
      || !description.includes('Match columns by gridColumnStart')) {
    fail(check, 'docx_read_table does not describe retained rows, continuation, and logical columns');
  }
  if (Object.hasOwn(input?.properties ?? {}, 'limit')
      || Object.hasOwn(input?.properties ?? {}, 'offset')
      || input?.properties?.continuation?.type !== 'string'
      || cell?.properties?.gridColumnStart?.type !== 'integer'
      || cell?.properties?.gridColumnStart?.minimum !== 0
      || cell?.properties?.text?.type !== 'string'
      || cell?.properties?.logicalText?.type !== 'string'
      || Object.hasOwn(cell?.properties ?? {}, 'paragraphs')
      || receipt?.properties?.nextContinuation?.type !== 'string'
      || receipt?.required?.includes('nextContinuation')
      || receipt?.properties?.retainedRowCount?.type !== 'integer'
      || receipt?.properties?.detailPageRetained?.type !== 'boolean') {
    fail(check, 'docx_read_table response is not a compact page backed by one detailed page');
  }
  note('DOCX table pages expose logical columns and compact cell text while retaining selected-page detail on disk');
}

function checkDocxTableIndexContract(tools) {
  const check = 'docx-table-index-contract';
  const tool = tools.find(entry => entry?.name === 'docx_table_index');
  const description = tool?.description || '';
  const input = tool?.inputSchema;
  const table = tool?.outputSchema?.properties?.tables?.items;
  if (Object.hasOwn(input?.properties ?? {}, 'limit')
      || input?.properties?.offset?.type !== 'integer'
      || !description.includes('provider chooses page size')
      || !description.includes('pass one returned address unchanged')) {
    fail(check, 'docx_table_index does not own bounded page sizing and native-address continuation');
  }
  const properties = table?.properties ?? {};
  const nullableString = schema => Array.isArray(schema?.anyOf)
    && schema.anyOf.some(option => option?.type === 'string')
    && schema.anyOf.some(option => option?.type === 'null');
  if (properties.address?.type !== 'object'
      || properties.rowCount?.type !== 'integer'
      || properties.columnCount?.type !== 'integer'
      || properties.textPreview?.type !== 'string'
      || !nullableString(properties.precedingText)
      || !nullableString(properties.followingText)
      || Object.hasOwn(properties, 'parentAddress')
      || Object.hasOwn(properties, 'textLength')
      || Object.hasOwn(properties, 'precedingParagraph')
      || Object.hasOwn(properties, 'followingParagraph')) {
    fail(check, 'docx_table_index response is not a compact native-address locator');
  }
  note('DOCX table index owns response page size and returns compact native-address locators');
}

function checkProviderOwnedReadPaging(tools) {
  const check = 'provider-owned-read-paging';
  const list = tools.find(entry => entry?.name === 'docx_list_objects');
  const input = list?.inputSchema;
  if (Object.hasOwn(input?.properties ?? {}, 'limit')
      || input?.properties?.offset?.minimum !== 0
      || !(list?.description || '').includes('provider chooses the bounded page size')) {
    fail(check, 'docx_list_objects does not publish provider-owned bounded paging');
  }
  note('all five offset-based readers keep technical page size inside the provider');
}

function collectWriteFileNodes(schema, location = '$', found = []) {
  if (Array.isArray(schema)) {
    schema.forEach((entry, index) => collectWriteFileNodes(entry, `${location}[${index}]`, found));
    return found;
  }
  if (!schema || typeof schema !== 'object') return found;
  if (schema[fileRoleKey] === 'write') found.push({ location, schema });
  for (const [key, value] of Object.entries(schema)) {
    collectWriteFileNodes(value, `${location}.${key}`, found);
  }
  return found;
}

function checkReadOnlyFileEffects(officeTools, textTools) {
  const check = 'read-only-file-effects';
  const readOnlyTools = [...officeTools, ...textTools]
    .filter(tool => tool?.annotations?.readOnlyHint === true);
  for (const tool of readOnlyTools) {
    for (const entry of collectWriteFileNodes(tool.inputSchema)) {
      if (entry.schema[fileEffectKey] !== false) {
        fail(check, `${tool.name} read evidence is not marked non-effect at ${entry.location}`);
      }
    }
  }
  note(`${readOnlyTools.length} read-only tools publish every evidence write as non-effect`);
}

function unboundedResponseArrays(schema, location = '$', found = []) {
  if (Array.isArray(schema)) {
    schema.forEach((entry, index) => unboundedResponseArrays(entry, `${location}[${index}]`, found));
    return found;
  }
  if (!schema || typeof schema !== 'object') return found;
  if (schema.type === 'array' && !Number.isInteger(schema.maxItems)) found.push(location);
  for (const [key, value] of Object.entries(schema)) {
    unboundedResponseArrays(value, `${location}.${key}`, found);
  }
  return found;
}

function checkBoundedInspectionOutputs(tools) {
  const check = 'bounded-inspection-output';
  const inspections = tools.filter(tool => tool?.name?.endsWith('_inspect'));
  for (const tool of inspections) {
    if (!tool.outputSchema || typeof tool.outputSchema !== 'object') {
      fail(check, `${tool.name} has no machine-readable output schema`);
      continue;
    }
    for (const location of unboundedResponseArrays(tool.outputSchema)) {
      fail(check, `${tool.name} exposes an unbounded response collection at ${location}`);
    }
  }
  note(`${inspections.length} inspect outputs keep complete evidence in artifacts and bound every response collection`);
}

function checkCompactInspectionSummaries(tools) {
  const check = 'compact-inspection-summaries';
  const expected = new Map([
    ['xlsx_inspect', 'sheets'],
    ['pptx_inspect', 'openingSlides'],
  ]);
  for (const [name, collection] of expected) {
    const schema = tools.find(tool => tool?.name === name)?.outputSchema;
    const summary = schema?.properties?.summary;
    const items = summary?.properties?.[collection];
    if (!schema?.required?.includes('summary')
        || summary?.type !== 'object'
        || items?.type !== 'array'
        || !Number.isInteger(items.maxItems)) {
      fail(check, `${name} does not publish a required bounded identity summary`);
    }
  }
  note('XLSX and PPTX inspection always return bounded identity summaries');
}

async function checkIdempotentReadArtifacts(tempRoot) {
  const check = 'idempotent-read-artifact';
  const output = path.join(tempRoot, 'idempotent-read.json');
  const payload = { schema: 'test/read-result', rows: [{ value: 'same' }] };
  const first = await writeIdempotentJsonArtifact(output, payload);
  const replay = await writeIdempotentJsonArtifact(output, payload);
  if (JSON.stringify(first) !== JSON.stringify(replay)) {
    fail(check, 'identical read replay did not return the existing artifact identity');
  }
  let rejected = false;
  try {
    await writeIdempotentJsonArtifact(output, { ...payload, rows: [{ value: 'different' }] });
  } catch (error) {
    rejected = error?.code === 'EEXIST';
  }
  if (!rejected || JSON.stringify(await readJson(output)) !== JSON.stringify(payload)) {
    fail(check, 'different read content did not preserve and reject the existing artifact');
  }
  note('read artifacts accept identical replay and reject different content without overwrite');
}

async function main() {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'tiwater-office-boundary-'));
  try {
    await checkDependencyGraph();
    await checkOfficeSourceOwnership();
    await checkTextSourceOwnership();
    await checkFixedRuntimeSurface();
    await checkPackageFiles();
    const { archive, manifest } = await packOfficePackage(tempRoot);
    const packageRoot = await extractArchive(archive, path.join(tempRoot, 'extracted'));
    await checkPackedPackage(manifest, packageRoot);
    await checkPublicSchemas(packageRoot);
    const { officeTools: toolNames, textTools } = await smokeInstalledPackage(archive, tempRoot);
    checkSourceBoundObservationOutputs(toolNames);
    checkLargeResultChannels(toolNames);
    checkXlsxRangeReadContract(toolNames);
    checkDocxCreateContract(toolNames);
    checkPptxBoundedReadContracts(toolNames);
    checkDocxMergedCellDescriptions(toolNames);
    checkDocxTableStreamingContract(toolNames);
    checkDocxTableIndexContract(toolNames);
    checkProviderOwnedReadPaging(toolNames);
    checkBoundedInspectionOutputs(toolNames);
    checkCompactInspectionSummaries(toolNames);
    await checkIdempotentReadArtifacts(tempRoot);
    const packedPackage = await readJson(path.join(packageRoot, 'package.json'));
    await checkGeneratedManifest(packageRoot, toolNames, packedPackage);
    await checkTextPublishedSurface(packageRoot, textTools, packedPackage);
    checkEvidenceRoleMetadata(toolNames, textTools);
    checkEffectKindMetadata(toolNames, textTools);
    checkReadOnlyFileEffects(toolNames, textTools);
  } catch (error) {
    fail('gate-runtime', error.stack || error.message);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }

  for (const message of notes) console.log(`PASS ${message}`);
  if (failures.length > 0) {
    for (const message of failures) console.error(`FAIL ${message}`);
    process.exitCode = 1;
    return;
  }
  console.log('PASS release boundary gate');
}

await main();
