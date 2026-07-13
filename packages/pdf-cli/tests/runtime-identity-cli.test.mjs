import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

import Ajv2020 from 'ajv/dist/2020.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const packageDir = path.resolve(here, '..');
const repoRoot = path.resolve(packageDir, '../..');
const python = process.env.TIWATER_PDF_TEST_PYTHON
  ?? path.join(packageDir, '.venv', 'bin', 'python');

const capabilitySchema = JSON.parse(fs.readFileSync(
  path.join(repoRoot, 'contracts/runtime/runtime-capabilities.schema.json'),
  'utf8',
));
const evidenceSchema = JSON.parse(fs.readFileSync(
  path.join(repoRoot, 'contracts/runtime/runtime-evidence-envelope.schema.json'),
  'utf8',
));
const ajv = new Ajv2020({ allErrors: true, strict: false });
const validateCapabilities = ajv.compile(capabilitySchema);
const validateEvidence = ajv.compile(evidenceSchema);

function runPdfCli(...args) {
  return spawnSync(python, ['-m', 'tiwater_pdf.cli', ...args], {
    cwd: repoRoot,
    encoding: 'utf8',
    env: {
      ...process.env,
      PYTHONPATH: packageDir,
      OPENAI_API_KEY: 'must-not-appear-in-runtime-identity',
      HTTPS_PROXY: 'http://127.0.0.1:1',
      HTTP_PROXY: 'http://127.0.0.1:1',
    },
  });
}

function assertJsonOnly(result, expectedExitCode) {
  assert.equal(result.status, expectedExitCode, result.stderr);
  assert.equal(result.signal, null);
  assert.equal(result.stderr, '');
  assert.doesNotMatch(result.stdout, /must-not-appear-in-runtime-identity/);
  return JSON.parse(result.stdout);
}

function assertSchemaValid(validate, value) {
  assert.equal(validate(value), true, ajv.errorsText(validate.errors));
}

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(
      Object.keys(value).sort().map((key) => [key, canonicalize(value[key])]),
    );
  }
  return value;
}

function assertCanonicalArtifact(evidence) {
  const bytes = Buffer.from(JSON.stringify(canonicalize(evidence.payload)));
  const sha256 = crypto.createHash('sha256').update(bytes).digest('hex');
  assert.deepEqual(evidence.artifact, {
    artifactId: `sha256:${sha256}`,
    sizeBytes: bytes.length,
    sha256,
    mediaType: 'application/json',
    encoding: 'canonical-json',
    schema: { id: 'tiwater.runtime.identify-payload', version: '1.0.0' },
  });
}

function createPdf(target, encrypted = false) {
  const script = String.raw`
import fitz
import sys

target, encrypted = sys.argv[1], sys.argv[2] == "true"
document = fitz.open()
page = document.new_page()
page.insert_text((72, 72), "runtime identity fixture")
options = {}
if encrypted:
    options = {
        "encryption": fitz.PDF_ENCRYPT_AES_256,
        "owner_pw": "owner-secret",
        "user_pw": "user-secret",
    }
document.save(target, **options)
document.close()
`;
  const result = spawnSync(python, ['-c', script, target, String(encrypted)], {
    encoding: 'utf8',
    env: { ...process.env, PYTHONPATH: packageDir },
  });
  assert.equal(result.status, 0, result.stderr);
}

test('capabilities --json emits the schema-valid PDF runtime descriptor', () => {
  const descriptor = assertJsonOnly(runPdfCli('capabilities', '--json'), 0);
  const pyproject = fs.readFileSync(path.join(packageDir, 'pyproject.toml'), 'utf8');
  const packageVersion = /^version = "([^"]+)"$/m.exec(pyproject)?.[1];

  assertSchemaValid(validateCapabilities, descriptor);
  assert.equal(descriptor.package.name, 'tiwater-pdf');
  assert.equal(descriptor.package.version, '0.16.0');
  assert.equal(descriptor.package.version, packageVersion);
  assert.deepEqual(descriptor.runtime, {
    family: 'pdf',
    name: 'tiwater-pdf',
    version: '0.16.0',
  });
  assert.deepEqual(descriptor.supportedKinds, [{
    fileKind: 'pdf',
    mediaTypes: ['application/pdf'],
    signatureKinds: ['pdf-header-pymupdf-open'],
  }]);
  assert.equal(descriptor.descriptorCommand.mutates, false);
  assert.equal(descriptor.identifyProbe.mutates, false);
});

test('identify supports renamed and encrypted PDFs from exact bytes deterministically', () => {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-runtime-supported-'));
  try {
    for (const [name, encrypted] of [['renamed.payload', false], ['encrypted.bin', true]]) {
      const sourcePath = path.join(temporary, name);
      createPdf(sourcePath, encrypted);
      const before = fs.readFileSync(sourcePath);

      const first = assertJsonOnly(runPdfCli('identify', sourcePath, '--json'), 0);
      const second = assertJsonOnly(runPdfCli('identify', sourcePath, '--json'), 0);

      assertSchemaValid(validateEvidence, first);
      assert.deepEqual(first, second);
      assert.equal(first.status, 'supported');
      assert.equal(first.file.fileKind, 'pdf');
      assert.equal(first.file.mediaType, 'application/pdf');
      assert.equal(first.file.signature.status, 'matched');
      assert.equal(first.file.signature.kind, 'pdf-header-pymupdf-open');
      assert.ok(first.file.signature.evidence.includes(`pdf:encrypted=${encrypted}`));
      assert.equal(first.payload.encrypted, encrypted);
      assert.equal(first.source.sizeBytes, before.length);
      assert.equal(first.source.sha256, crypto.createHash('sha256').update(before).digest('hex'));
      assertCanonicalArtifact(first);
      assert.deepEqual(fs.readFileSync(sourcePath), before);
    }
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
});

test('identify types invalid content as unsupported and source read errors as failed', () => {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-runtime-rejected-'));
  try {
    const unsupported = new Map([
      ['fake.pdf', Buffer.from('not a PDF')],
      ['image.bin', Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl8sAAAAASUVORK5CYII=',
        'base64',
      )],
      ['header-only.pdf', Buffer.from('%PDF-1.7\n')],
      ['truncated.pdf', Buffer.from('%PDF-1.7\n1 0 obj\n<< /Type /Catalog >>')],
    ]);
    const invalidVersionPath = path.join(temporary, 'invalid-version-source.pdf');
    createPdf(invalidVersionPath);
    const invalidVersionBytes = fs.readFileSync(invalidVersionPath);
    invalidVersionBytes.write('%PDF-1.9', 0, 'ascii');
    unsupported.set('invalid-version.pdf', invalidVersionBytes);
    for (const [name, bytes] of unsupported) {
      const sourcePath = path.join(temporary, name);
      fs.writeFileSync(sourcePath, bytes);
      const evidence = assertJsonOnly(runPdfCli('identify', sourcePath, '--json'), 0);
      assertSchemaValid(validateEvidence, evidence);
      assert.equal(evidence.status, 'unsupported', name);
      assert.equal(evidence.source.sha256, crypto.createHash('sha256').update(bytes).digest('hex'));
      assert.equal(evidence.file.fileKind, null);
      assert.equal(evidence.file.mediaType, null);
      assertCanonicalArtifact(evidence);
    }

    const missingPath = path.join(temporary, 'missing.pdf');
    const directoryPath = path.join(temporary, 'directory.pdf');
    fs.mkdirSync(directoryPath);
    const unreadablePath = path.join(temporary, 'unreadable.pdf');
    fs.writeFileSync(unreadablePath, '%PDF-1.7\n');
    fs.chmodSync(unreadablePath, 0o000);
    try {
      for (const sourcePath of [missingPath, directoryPath, unreadablePath]) {
        const evidence = assertJsonOnly(runPdfCli('identify', sourcePath, '--json'), 1);
        assertSchemaValid(validateEvidence, evidence);
        assert.equal(evidence.status, 'failed');
        assert.equal(evidence.failureStage, 'source-read');
        assert.equal(evidence.source, null);
        assert.equal(evidence.file.signature.status, 'not-checked');
        assert.deepEqual(evidence.file.signature.evidence, []);
        assert.deepEqual(evidence.errors, [{
          code: 'source-read-failed',
          message: 'The source bytes could not be read.',
        }]);
        assertCanonicalArtifact(evidence);
      }
    } finally {
      fs.chmodSync(unreadablePath, 0o600);
    }
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
});
