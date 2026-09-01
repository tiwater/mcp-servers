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

const execFileAsync = promisify(execFile);
const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, '..');
const serverRoot = path.join(repoRoot, 'servers');
const serverPackagePath = path.join(serverRoot, 'package.json');
const serverLockPath = path.join(serverRoot, 'package-lock.json');
const generatedManifestRelativePath = 'office/contracts/tiwater-office-provider-contract-manifest-v1.json';
const generatedContractDeclaration = 'office/contracts/*.schema.json';
const fileRoleKey = 'x-tiwater-file-role';
const fileEffectKey = 'x-tiwater-file-effect';
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
  'office/index.mjs',
  'office/README.md',
  generatedManifestRelativePath,
];
const requiredPackageDeclarations = [
  generatedManifestRelativePath,
  generatedContractDeclaration,
];
const providerContractRoots = [
  'packages/convert-cli/schemas',
  'packages/docx-cli/contracts',
  'packages/pptx-cli/contracts',
  'packages/xlsx-cli/contracts',
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

async function checkFixedRuntimeSurface() {
  const check = 'fixed-runtime-surface';
  const officeSource = await readFile(path.join(serverRoot, 'office', 'index.mjs'), 'utf8');
  const fixedNames = new Set([
    ...[...officeSource.matchAll(/\{"name":"((?:docx|xlsx)_[^"]+)"/g)].map(match => match[1]),
    ...[...officeSource.matchAll(/fixedEdit\('((?:docx|xlsx|pptx)_[^']+)'/g)].map(match => match[1]),
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
  if (!packedPaths.has(generatedManifestRelativePath)) {
    fail(check, `generated provider contract manifest is absent from pack: ${generatedManifestRelativePath}`);
  }
  if (contractPaths.length === 0) {
    fail(check, 'Office MCP pack must contain at least one generated public contract schema');
  } else {
    note(`Office MCP pack contains ${contractPaths.length} public contract schemas`);
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
    .filter(file => file.endsWith('.schema.json') && file.includes(`${path.sep}contracts${path.sep}`))
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
  function visit(node, location, propertyName = '') {
    if (!node || typeof node !== 'object' || Array.isArray(node)) return;
    const declaredRole = node[fileRoleKey];
    const declaredEffect = node[fileEffectKey];
    if (declaredRole !== undefined) {
      if (node.type !== 'string' || !['read', 'write'].includes(declaredRole)) {
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
      visit(child, `${location}.properties.${name}`, name);
    }
    if (node.items) visit(node.items, `${location}.items`);
    for (const keyword of ['allOf', 'anyOf', 'oneOf']) {
      for (const [index, child] of (node[keyword] || []).entries()) {
        visit(child, `${location}.${keyword}[${index}]`, propertyName);
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

async function smokeInstalledPackage(archive, tempRoot) {
  const check = 'isolated-smoke';
  const installRoot = path.join(tempRoot, 'unrelated-install');
  await mkdirIfMissing(installRoot);
  if (!path.relative(repoRoot, installRoot).startsWith('..')) {
    fail(check, `smoke directory is inside the Lucid/provider repository: ${installRoot}`);
    return;
  }
  await execFileAsync('npm', [
    'install', '--ignore-scripts', '--no-audit', '--no-fund', '--package-lock=false', '--prefix', installRoot, archive,
  ], { cwd: installRoot, maxBuffer: 8 * 1024 * 1024 });

  const executable = path.join(installRoot, 'node_modules', '.bin', 'tiwater-office-mcp');
  if (!(await exists(executable))) {
    fail(check, 'installed package did not expose tiwater-office-mcp executable');
    return;
  }

  const response = await new Promise((resolve, reject) => {
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
  if (!response.initialized) fail(check, 'MCP initialize did not complete');
  if (!response.serverInstructions.includes('A read-only output path is an immutable artifact identity')
      || !response.serverInstructions.includes('an identical request may replay it')
      || !response.serverInstructions.includes('every different request uses a different path')) {
    fail(check, 'MCP instructions do not publish immutable read artifact path semantics');
  }
  note(`isolated MCP initialize and tools/list completed (${response.tools?.length || 0} tools)${response.stderr.trim() ? ' with stderr output' : ''}`);
  return response.tools || [];
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
    'docx_validate_toc_style_policy', 'xlsx_inspect', 'xlsx_export_json', 'xlsx_validate',
    'pptx_inspect', 'pptx_export_json', 'pptx_validate',
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
    'docx_validate_toc_style_policy', 'xlsx_inspect', 'xlsx_export_json', 'xlsx_validate',
    'pptx_inspect', 'pptx_export_json', 'pptx_validate',
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
  const setBody = tools.find(tool => tool?.name === 'docx_set_table_body');
  const setBodyDescription = setBody?.description || '';
  const setBodyCell = setBody?.inputSchema?.properties?.rows?.items?.properties?.cells?.items;
  if (!readDescription.includes('restart')
      || !readDescription.includes('continue cell is not an independent row value')
      || !readDescription.includes('logicalText resolves the restart cell value')) {
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
  if (!setBodyDescription.includes('cover every row completely with explicit cells')
      || !setBodyDescription.includes('restart cell followed by continue cells')
      || setBodyCell?.properties?.verticalMerge?.type !== 'string'
      || !setBodyCell?.properties?.verticalMerge?.enum?.includes('restart')
      || !setBodyCell?.properties?.verticalMerge?.enum?.includes('continue')
      || Object.hasOwn(setBodyCell?.properties ?? {}, 'rowSpan')) {
    fail(check, 'docx_set_table_body does not expose explicit native vertical-merge cells');
  }
  note('DOCX table read/write descriptions preserve vertical-merge logical-cell semantics');
}

function checkDocxTableStreamingContract(tools) {
  const check = 'docx-table-streaming-contract';
  const tool = tools.find(entry => entry?.name === 'docx_read_table');
  const description = tool?.description || '';
  const row = tool?.outputSchema?.properties?.rows?.items;
  const cell = row?.properties?.cells?.items;
  const receipt = tool?.outputSchema?.properties?.receipt;
  if (!description.includes('selected row page')
      || !description.includes('never builds another whole-table data object')
      || !description.includes('zero-based logical gridColumnStart')
      || !description.includes('Match columns across rows by gridColumnStart')
      || !description.includes('physical cell ordinal and is not a column identity')
      || !description.includes('receipt.remaining is navigation information, not an obligation')
      || !description.includes('receipt.nextOffset is present only when another row page exists')
      || !description.includes('blank template rows need not be paged through')) {
    fail(check, 'docx_read_table does not describe logical columns and current-decision page consumption');
  }
  if (cell?.properties?.gridColumnStart?.type !== 'integer'
      || cell?.properties?.gridColumnStart?.minimum !== 0
      || cell?.properties?.text?.type !== 'string'
      || cell?.properties?.logicalText?.type !== 'string'
      || Object.hasOwn(cell?.properties ?? {}, 'paragraphs')
      || receipt?.properties?.nextOffset?.type !== 'integer'
      || receipt?.required?.includes('nextOffset')
      || receipt?.properties?.detailPageRetained?.type !== 'boolean') {
    fail(check, 'docx_read_table response is not a compact page backed by one detailed page');
  }
  note('DOCX table pages expose logical columns and compact cell text while retaining selected-page detail on disk');
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
    await checkFixedRuntimeSurface();
    await checkPackageFiles();
    const { archive, manifest } = await packOfficePackage(tempRoot);
    const packageRoot = await extractArchive(archive, path.join(tempRoot, 'extracted'));
    await checkPackedPackage(manifest, packageRoot);
    await checkPublicSchemas(packageRoot);
    const toolNames = await smokeInstalledPackage(archive, tempRoot);
    checkSourceBoundObservationOutputs(toolNames);
    checkLargeResultChannels(toolNames);
    checkDocxMergedCellDescriptions(toolNames);
    checkDocxTableStreamingContract(toolNames);
    checkBoundedInspectionOutputs(toolNames);
    checkCompactInspectionSummaries(toolNames);
    await checkIdempotentReadArtifacts(tempRoot);
    const packedPackage = await readJson(path.join(packageRoot, 'package.json'));
    await checkGeneratedManifest(packageRoot, toolNames, packedPackage);
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
