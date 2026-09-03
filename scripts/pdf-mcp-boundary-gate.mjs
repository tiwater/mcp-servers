#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { access, mkdtemp, mkdir, readFile, rm, stat } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import readline from 'node:readline';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { promisify } from 'node:util';
import {
  assertEvidenceToolContract,
  evidenceRoleMetadataKey,
} from '../servers/_shared/evidence-role.mjs';
import { effectKindMetadataKey } from '../servers/_shared/effect-kind.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packageRoot = path.join(repoRoot, 'servers');
const pdfRoot = path.join(packageRoot, 'pdf');
const expectedTools = [
  'pdf_extract_table_details',
  'pdf_extract_tables',
  'pdf_find_table',
  'pdf_inspect',
  'pdf_ocr',
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
  try { await access(file); return true; } catch { return false; }
}

async function json(file) {
  return JSON.parse(await readFile(file, 'utf8'));
}

async function sha256(file) {
  return createHash('sha256').update(await readFile(file)).digest('hex');
}

async function packageIdentity() {
  const [packageJson, lock] = await Promise.all([
    json(path.join(packageRoot, 'package.json')),
    json(path.join(packageRoot, 'package-lock.json')),
  ]);
  if (packageJson.name !== '@tiwater/office-mcp' || packageJson.private === true
    || packageJson.bin?.['tiwater-office-mcp'] !== 'office/index.mjs'
    || packageJson.bin?.['tiwater-pdf-mcp'] !== 'pdf/index.mjs') {
    fail('package-identity', 'shared distribution identity or independent executable bindings are invalid');
  }
  if (lock.packages?.['']?.name !== packageJson.name
    || lock.packages?.['']?.version !== packageJson.version) {
    fail('package-identity', 'package-lock root identity does not match package.json');
  }
  note(`shared MCP distribution identity ${packageJson.name}@${packageJson.version}`);
  return packageJson;
}

async function contractIdentity(packageJson) {
  const manifestPath = path.join(
    pdfRoot, 'contracts', 'tiwater-pdf-provider-contract-manifest-v1.json',
  );
  const manifest = await json(manifestPath);
  if (manifest.schema !== 'tiwater.pdf-provider-contract-manifest/v1'
    || manifest.provider?.id !== packageJson.name
    || manifest.provider?.version !== packageJson.version
    || manifest.runtime?.command !== 'tiwater-pdf') {
    fail('provider-contracts', 'manifest identity does not match package or fixed runtime');
  }
  const names = (manifest.tools || []).map(entry => entry.name).sort();
  if (JSON.stringify(names) !== JSON.stringify(expectedTools)) {
    fail('provider-contracts', `unexpected contract tool set: ${names.join(',')}`);
  }
  for (const entry of manifest.tools || []) {
    const contractPath = path.join(pdfRoot, entry.inputContract?.path || '');
    if (!contractPath.startsWith(path.join(pdfRoot, 'contracts') + path.sep)
      || !await exists(contractPath)
      || await sha256(contractPath) !== entry.inputContract?.sha256) {
      fail('provider-contracts', `invalid input contract binding for ${entry.name}`);
      continue;
    }
    const schema = await json(contractPath);
    if (schema.type !== 'object'
      || schema.properties?.input?.['x-tiwater-file-role'] !== 'read'
      || schema.properties?.output?.['x-tiwater-file-role'] !== 'write'
      || schema.properties?.output?.['x-tiwater-file-effect'] !== false) {
      fail('provider-contracts', `${entry.name} lacks published file role/effect metadata`);
    }
  }
  note(`provider manifest binds ${names.length} PDF input contracts`);
}

async function pack(tempRoot) {
  const destination = path.join(tempRoot, 'pack');
  await mkdir(destination, { recursive: true });
  const { stdout } = await execFileAsync('npm', [
    'pack', '--json', '--ignore-scripts', '--pack-destination', destination,
  ], { cwd: packageRoot, maxBuffer: 4 * 1024 * 1024 });
  const values = JSON.parse(stdout);
  if (!Array.isArray(values) || values.length !== 1) throw new Error('npm-pack-result-invalid');
  const manifest = values[0];
  const paths = new Set(manifest.files.map(entry => entry.path));
  const required = [
    'package.json', 'office/index.mjs', 'pdf/index.mjs', 'pdf/README.md',
    '_shared/mcp-stdio.mjs', '_shared/tool-runtime.mjs', '_shared/large-json-result.mjs',
    '_shared/evidence-role.mjs',
    '_shared/effect-kind.mjs',
    'office/contracts/tiwater-office-provider-contract-manifest-v1.json',
    'pdf/contracts/tiwater-pdf-provider-contract-manifest-v1.json',
  ];
  if (manifest.name !== '@tiwater/office-mcp' || required.some(file => !paths.has(file))) {
    fail('pack-manifest', 'packed shared distribution is missing an MCP executable or contract');
  }
  note(`npm pack contains ${manifest.entryCount} entries and both independent MCP surfaces`);
  return path.join(destination, manifest.filename);
}

async function createPdf(file) {
  const program = [
    'import fitz, sys',
    'document = fitz.open()',
    'page = document.new_page(width=612, height=792)',
    'page.insert_text((72, 96), "Quarterly PDF observation")',
    'document.set_metadata({"title": "Quarterly Report", "author": "Tiwater"})',
    'document.save(sys.argv[1])',
  ].join('; ');
  await execFileAsync('python3', ['-c', program, file]);
  if ((await stat(file)).size < 1) throw new Error('real-pdf-generation-failed');
}

function startClient(executable, cwd, environment) {
  const child = spawn(executable, [], { cwd, env: environment, stdio: ['pipe', 'pipe', 'pipe'] });
  const pending = new Map();
  let stderr = '';
  readline.createInterface({ input: child.stdout, crlfDelay: Infinity }).on('line', line => {
    let message;
    try { message = JSON.parse(line); } catch { return; }
    const request = pending.get(message.id);
    if (request) {
      pending.delete(message.id);
      clearTimeout(request.timer);
      request.resolve(message);
    }
  });
  child.stderr.on('data', chunk => { stderr += chunk.toString('utf8'); });
  child.on('exit', (code, signal) => {
    for (const request of pending.values()) {
      clearTimeout(request.timer);
      request.reject(new Error(`MCP exited code=${code} signal=${signal}: ${stderr}`));
    }
    pending.clear();
  });
  return {
    child,
    request(id, method, params = {}) {
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          pending.delete(id);
          reject(new Error(`MCP request timeout: ${method}`));
        }, 30_000);
        pending.set(id, { resolve, reject, timer });
        child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
      });
    },
  };
}

async function isolatedSmoke(archive, tempRoot, packageJson) {
  const initialFailureCount = failures.length;
  const installRoot = path.join(tempRoot, 'clean-install');
  await mkdir(installRoot, { recursive: true });
  await execFileAsync('npm', [
    'install', '--ignore-scripts', '--no-audit', '--no-fund', '--package-lock=false',
    '--prefix', installRoot, archive,
  ], { cwd: installRoot, maxBuffer: 8 * 1024 * 1024 });
  const binRoot = path.join(installRoot, 'node_modules', '.bin');
  const pdfExecutable = path.join(binRoot, 'tiwater-pdf-mcp');
  const officeExecutable = path.join(binRoot, 'tiwater-office-mcp');
  if (!await exists(pdfExecutable) || !await exists(officeExecutable)) {
    throw new Error('clean-install-executable-missing');
  }

  const environment = {
    ...process.env,
    PATH: `${binRoot}${path.delimiter}${process.env.PATH || ''}`,
  };
  const officeClient = startClient(officeExecutable, installRoot, environment);
  try {
    const initialized = await officeClient.request(101, 'initialize', {
      protocolVersion: '2025-06-18',
      capabilities: {},
      clientInfo: { name: 'pdf-release-boundary-gate', version: '1.0.0' },
    });
    const listed = await officeClient.request(102, 'tools/list');
    const officeTools = listed.result?.tools;
    if (initialized.error
      || initialized.result?.serverInfo?.name !== 'tiwater-office'
      || initialized.result?.serverInfo?.version !== packageJson.version
      || !Array.isArray(officeTools) || officeTools.length === 0
      || officeTools.some(tool => tool?.name?.startsWith('pdf_'))) {
      fail('isolated-smoke', 'Office executable identity or Office-only tool surface is invalid');
    }
  } finally {
    officeClient.child.kill('SIGTERM');
  }

  const input = path.join(tempRoot, 'current.pdf');
  const output = path.join(tempRoot, 'pdf-inspection.json');
  await createPdf(input);
  const client = startClient(pdfExecutable, installRoot, environment);
  try {
    const initialized = await client.request(1, 'initialize', {
      protocolVersion: '2025-06-18',
      capabilities: {},
      clientInfo: { name: 'pdf-release-boundary-gate', version: '1.0.0' },
    });
    if (initialized.error
      || initialized.result?.serverInfo?.name !== 'tiwater-pdf'
      || initialized.result?.serverInfo?.version !== packageJson.version) {
      fail('isolated-smoke', 'provider initialize identity/version mismatch');
    }
    const listed = await client.request(2, 'tools/list');
    const tools = listed.result?.tools;
    const names = Array.isArray(tools) ? tools.map(tool => tool.name).sort() : [];
    if (listed.error || JSON.stringify(names) !== JSON.stringify(expectedTools)) {
      fail('isolated-smoke', `clean tools/list mismatch: ${names.join(',')}`);
    }
    for (const tool of tools || []) {
      if (tool.annotations?.readOnlyHint !== true
        || tool.annotations?.idempotentHint !== true
        || tool.annotations?.destructiveHint !== false
        || tool.annotations?.openWorldHint !== false) {
        fail('isolated-smoke', `${tool.name} annotations are incomplete`);
      }
    }
    for (const tool of tools || []) {
      if (tool?._meta?.[effectKindMetadataKey] !== undefined) {
        fail('isolated-smoke', `${tool.name} publishes an Office effect kind`);
      }
      const hasMetadata = tool?._meta?.[evidenceRoleMetadataKey] !== undefined;
      if (tool.name !== 'pdf_inspect') {
        if (hasMetadata) fail('isolated-smoke', `${tool.name} publishes an unexpected evidence role`);
        continue;
      }
      try {
        assertEvidenceToolContract(tool, 'document-observation');
      } catch (error) {
        fail('isolated-smoke', error.message);
      }
    }
    const inspectTool = (tools || []).find(tool => tool.name === 'pdf_inspect');
    if (!inspectTool?.inputSchema?.required?.includes('output')
      || !inspectTool?.outputSchema?.required?.includes('identity')
      || inspectTool?.outputSchema?.properties?.identity?.properties?.openingPages?.maxItems !== 3) {
      fail('isolated-smoke', 'pdf_inspect does not publish durable output and bounded identity');
    }
    const inspected = await client.request(3, 'tools/call', {
      name: 'pdf_inspect',
      arguments: { input, output, returnContent: false },
    });
    const result = inspected.result?.structuredContent;
    if (inspected.error || inspected.result?.isError === true
      || result?.tool !== 'pdf_inspect'
      || result?.identity?.format !== 'pdf'
      || result?.identity?.pageCount !== 1
      || result?.identity?.title !== 'Quarterly Report'
      || result?.identity?.openingPages?.length !== 1
      || result?.artifact?.path !== output
      || result?.artifact?.sha256 !== await sha256(output)
      || result?.sources?.[0]?.path !== input
      || result?.sources?.[0]?.sha256 !== await sha256(input)
      || result?.receipt?.contentWritten !== true
      || result?.receipt?.contentReturned !== false) {
      fail('isolated-smoke', `real PDF inspection is not bound to compact identity and durable artifacts: ${JSON.stringify(inspected)}`);
    }
    const retained = await exists(output) ? await json(output) : null;
    const retainedDocument = retained?.document ?? retained;
    if (retainedDocument?.pages !== 1 || retainedDocument.metadata?.title !== 'Quarterly Report') {
      fail('isolated-smoke', 'durable PDF observation is incomplete');
    }
    const replayed = await client.request(4, 'tools/call', {
      name: 'pdf_inspect',
      arguments: { input, output, returnContent: false },
    });
    if (replayed.error
      || replayed.result?.structuredContent?.artifact?.sha256 !== result?.artifact?.sha256
      || await sha256(output) !== result?.artifact?.sha256) {
      fail('isolated-smoke', 'identical pdf_inspect replay did not retain identical artifact identity');
    }
    if (failures.length === initialFailureCount) {
      note('clean npm install kept Office/PDF tools isolated and completed real pdf_inspect with idempotent replay');
    }
  } finally {
    client.child.kill('SIGTERM');
  }
}

async function main() {
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), 'tiwater-pdf-mcp-boundary-'));
  try {
    const packageJson = await packageIdentity();
    await contractIdentity(packageJson);
    const archive = await pack(tempRoot);
    await isolatedSmoke(archive, tempRoot, packageJson);
  } catch (error) {
    fail('gate-runtime', error.stack || error.message);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
  for (const message of notes) console.log(`PASS ${message}`);
  if (failures.length > 0) {
    for (const message of failures) console.error(`FAIL ${message}`);
    process.exitCode = 1;
  } else {
    console.log('PASS PDF MCP release boundary gate');
  }
}

await main();
