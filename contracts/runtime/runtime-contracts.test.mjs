import assert from 'node:assert/strict';
import Ajv2020 from 'ajv/dist/2020.js';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';

const here = path.dirname(fileURLToPath(import.meta.url));
const fixturePath = (name) => path.join(here, 'fixtures', name);
const schemaPath = (name) => path.join(here, name);
const readJson = (file) => JSON.parse(fs.readFileSync(file, 'utf8'));
const fixture = (name) => readJson(fixturePath(name));
const schemaNames = [
  'runtime-capabilities.schema.json',
  'runtime-evidence-envelope.schema.json',
  'edit-report.schema.json',
];
const schemas = Object.fromEntries(schemaNames.map((name) => [name, readJson(schemaPath(name))]));
const ajv = new Ajv2020({ allErrors: true, strict: false });
const schemaValidators = Object.fromEntries(
  Object.entries(schemas).map(([name, schema]) => [name, ajv.compile(schema)]),
);

function resolvePointer(document, pointer) {
  assert.ok(pointer.startsWith('#/'), `only local schema refs are allowed: ${pointer}`);
  return pointer.slice(2).split('/').reduce((current, segment) => {
    const key = segment.replaceAll('~1', '/').replaceAll('~0', '~');
    assert.ok(current && Object.hasOwn(current, key), `unresolved schema ref: ${pointer}`);
    return current[key];
  }, document);
}

function assertLocalRefsResolve(value, document) {
  if (Array.isArray(value)) {
    for (const item of value) assertLocalRefsResolve(item, document);
    return;
  }
  if (!value || typeof value !== 'object') return;
  if (typeof value.$ref === 'string') resolvePointer(document, value.$ref);
  for (const child of Object.values(value)) assertLocalRefsResolve(child, document);
}

function canonicalize(value) {
  if (Array.isArray(value)) return value.map(canonicalize);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonicalize(value[key])]));
  }
  if (typeof value === 'number' && !Number.isSafeInteger(value)) {
    throw new Error('canonical JSON v1 accepts safe integer numbers only');
  }
  return value;
}

function canonicalBytes(value) {
  return Buffer.from(JSON.stringify(canonicalize(value)));
}

function canonicalBytesFromRaw(rawJson) {
  assertCanonicalNumberLexemes(rawJson);
  return canonicalBytes(JSON.parse(rawJson));
}

function assertCanonicalNumberLexemes(rawJson) {
  let inString = false;
  let escaped = false;
  for (let index = 0; index < rawJson.length; index += 1) {
    const character = rawJson[index];
    if (inString) {
      if (escaped) escaped = false;
      else if (character === '\\') escaped = true;
      else if (character === '"') inString = false;
      continue;
    }
    if (character === '"') {
      inString = true;
      continue;
    }
    if (character !== '-' && (character < '0' || character > '9')) continue;

    const match = rawJson.slice(index).match(/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/);
    if (!match) continue;
    const token = match[0];
    if (/[.eE]/.test(token)) throw new Error(`canonical JSON requires a lexical integer: ${token}`);
    const integer = BigInt(token);
    if (integer < -9007199254740991n || integer > 9007199254740991n) {
      throw new Error(`canonical JSON requires a safe integer: ${token}`);
    }
    index += token.length - 1;
  }
}

function assertSchemaAccepts(schemaName, instance, label) {
  const validate = schemaValidators[schemaName];
  assert.equal(validate(instance), true, `${label}: ${ajv.errorsText(validate.errors)}`);
}

function assertSchemaRejects(schemaName, instance, label) {
  const validate = schemaValidators[schemaName];
  assert.equal(validate(instance), false, `${label} unexpectedly passed`);
}

function requireIdentity(value, label) {
  assert.equal(typeof value?.path, 'string', `${label}.path`);
  assert.ok(value.path.length > 0, `${label}.path`);
  assert.equal(Number.isSafeInteger(value?.sizeBytes), true, `${label}.sizeBytes`);
  assert.ok(value.sizeBytes >= 0, `${label}.sizeBytes non-negative`);
  assert.match(value?.sha256, /^[0-9a-f]{64}$/, `${label}.sha256`);
  assert.equal(value?.contentId, `sha256:${value.sha256}`, `${label}.contentId`);
}

function requireArtifact(artifact, payload, label) {
  const bytes = canonicalBytes(payload);
  const digest = createHash('sha256').update(bytes).digest('hex');
  assert.equal(artifact?.encoding, 'canonical-json', `${label}.encoding`);
  assert.equal(artifact?.sizeBytes, bytes.byteLength, `${label}.sizeBytes`);
  assert.equal(artifact?.sha256, digest, `${label}.sha256`);
  assert.equal(artifact?.artifactId, `sha256:${digest}`, `${label}.artifactId`);
  assert.ok(artifact?.schema?.id, `${label}.schema.id`);
  assert.ok(artifact?.schema?.version, `${label}.schema.version`);
}

function validateCapabilities(descriptor) {
  assert.equal(descriptor.schemaVersion, '1.0.0');
  assert.equal(descriptor.descriptorType, 'runtime-capabilities');
  for (const label of ['package', 'runtime', 'evidenceSchema']) {
    assert.ok(descriptor[label]?.name || descriptor[label]?.id, label);
    assert.ok(descriptor[label]?.version, `${label}.version`);
  }
  assert.deepEqual(descriptor.descriptorCommand, {
    command: 'capabilities',
    arguments: ['--json'],
    mutates: false,
  });
  assert.deepEqual(descriptor.identifyProbe?.outcomes, ['supported', 'unsupported', 'failed']);
  assert.equal(descriptor.identifyProbe?.mutates, false);
  assert.equal(descriptor.identifyProbe?.command, 'identify');
  assert.ok(descriptor.supportedKinds.length > 0);
  assert.ok(descriptor.commands.some((command) => command.name === 'capabilities' && command.mutates === false));
  assert.ok(descriptor.commands.some((command) => command.name === 'identify' && command.mutates === false));
  assert.equal(descriptor.identityPolicy?.nativeIds, 'runtime-native-only');
  assert.equal(descriptor.identityPolicy?.derivedIds, 'deterministic-and-explicit');
}

function validateEvidence(envelope) {
  assert.equal(envelope.schemaVersion, '1.0.0');
  assert.equal(envelope.envelopeType, 'runtime-evidence');
  assert.ok(['supported', 'unsupported', 'failed'].includes(envelope.status));
  const sourceReadFailure = envelope.status === 'failed' && envelope.failureStage === 'source-read';
  if (sourceReadFailure) assert.equal(envelope.source, null, 'source-read failure must not invent source identity');
  else requireIdentity(envelope.source, 'source');
  requireArtifact(envelope.artifact, envelope.payload, 'artifact');

  if (envelope.status === 'supported') {
    assert.equal(envelope.failureStage, null);
    assert.ok(envelope.file.fileKind);
    assert.ok(envelope.file.mediaType);
    assert.equal(envelope.file.signature.status, 'matched');
    assert.ok(envelope.file.signature.evidence.length > 0, 'supported signature evidence');
    assert.deepEqual(envelope.errors, []);
  } else if (envelope.status === 'unsupported') {
    assert.equal(envelope.failureStage, null);
    assert.equal(envelope.file.fileKind, null);
    assert.equal(envelope.file.mediaType, null);
    assert.notEqual(envelope.file.signature.status, 'matched');
    assert.deepEqual(envelope.objects, []);
    assert.deepEqual(envelope.errors, []);
  } else {
    assert.ok(envelope.failureStage, 'failed evidence requires failure stage');
    assert.equal(envelope.file.fileKind, null, 'failed file evidence cannot claim a kind');
    assert.equal(envelope.file.mediaType, null, 'failed file evidence cannot claim a media type');
    assert.notEqual(envelope.file.signature.status, 'matched', 'failed file evidence cannot claim a matched signature');
    assert.ok(envelope.errors.length > 0);
    assert.deepEqual(envelope.objects, []);
    if (sourceReadFailure) {
      assert.equal(envelope.file.signature.status, 'not-checked');
      assert.deepEqual(envelope.file.signature.evidence, []);
    }
  }

  const ids = new Set(envelope.objects.map((object) => object.objectId));
  assert.equal(ids.size, envelope.objects.length, 'object ids must be unique');
  for (const object of envelope.objects) {
    if (object.root) assert.equal(object.parentObjectId, null, `${object.objectId} root parent`);
    else assert.ok(ids.has(object.parentObjectId), `${object.objectId} parent must exist`);
    if (object.identity.kind === 'native') {
      assert.ok(object.identity.nativeId, `${object.objectId} nativeId`);
      assert.equal('derivation' in object.identity, false, `${object.objectId} must not claim a derivation`);
    } else {
      assert.equal(object.identity.kind, 'derived');
      assert.equal('nativeId' in object.identity, false, `${object.objectId} must not fabricate nativeId`);
      assert.ok(object.identity.derivation, `${object.objectId} derivation`);
      assert.ok(object.identity.inputs.length > 0, `${object.objectId} derivation inputs`);
    }
  }

  const byId = new Map(envelope.objects.map((object) => [object.objectId, object]));
  for (const start of envelope.objects) {
    const visited = new Set();
    let current = start;
    while (!current.root) {
      assert.equal(visited.has(current.objectId), false, `containment cycle at ${current.objectId}`);
      visited.add(current.objectId);
      current = byId.get(current.parentObjectId);
    }
  }
}

function validateEditReport(report, authoritativeRequest) {
  assert.equal(report.schemaVersion, '1.0.0');
  assert.equal(report.reportType, 'runtime-edit-report');
  requireIdentity(report.source, 'source');
  requireIdentity(report.output, 'output');
  assert.ok(Array.isArray(authoritativeRequest?.operations), 'authoritative request operations');
  requireArtifact(report.requestArtifact, authoritativeRequest, 'requestArtifact');
  assert.equal(report.operations.length, authoritativeRequest.operations.length, 'authoritative request count');
  assert.deepEqual(report.operations.map((entry) => entry.index), report.operations.map((_, index) => index));

  const counts = { applied: 0, noop: 0, rejected: 0, failed: 0 };
  for (const entry of report.operations) {
    assert.deepEqual(
      entry.requestedPayload,
      authoritativeRequest.operations[entry.index],
      `operation ${entry.index} must match authoritative request`,
    );
    assert.equal(entry.type, entry.requestedPayload.type, `operation ${entry.index} type`);
    assert.ok(Object.hasOwn(entry, 'appliedPayload'), `operation ${entry.index} complete applied payload`);
    assert.ok(Array.isArray(entry.targets), `operation ${entry.index} targets`);
    assert.ok(Array.isArray(entry.warnings), `operation ${entry.index} warnings`);
    assert.ok(Array.isArray(entry.errors), `operation ${entry.index} errors`);
    assert.ok(Object.hasOwn(counts, entry.status), `operation ${entry.index} status`);
    counts[entry.status] += 1;
    if (entry.status === 'rejected' || entry.status === 'failed') {
      assert.equal(entry.appliedPayload, null, `operation ${entry.index} rejected payload`);
      assert.ok(entry.errors.length > 0, `operation ${entry.index} errors required`);
    } else {
      assert.notEqual(entry.appliedPayload, null, `operation ${entry.index} applied payload required`);
    }
  }
  assert.deepEqual(report.summary, { requested: report.operations.length, ...counts });
}

test('publishes the three versioned runtime contract schemas', () => {
  for (const name of schemaNames) {
    assert.equal(fs.existsSync(schemaPath(name)), true, `${name} must exist`);
    const schema = readJson(schemaPath(name));
    assert.equal(schema.$schema, 'https://json-schema.org/draft/2020-12/schema');
    assert.ok(schema.$id.endsWith(name));
    assert.equal(schema.additionalProperties, false);
    assertLocalRefsResolve(schema, schema);
  }
});

test('Ajv 2020 compiles schemas and validates every positive contract fixture', () => {
  assertSchemaAccepts('runtime-capabilities.schema.json', fixture('runtime-capabilities.json'), 'capabilities');
  for (const name of [
    'supported-identify.json',
    'unsupported-identify.json',
    'failed-identify.json',
    'failed-source-read-identify.json',
  ]) {
    assertSchemaAccepts('runtime-evidence-envelope.schema.json', fixture(name), name);
  }
  assertSchemaAccepts('edit-report.schema.json', fixture('edit-report.json'), 'edit report');
});

test('Ajv rejects nested shape, version, uniqueness, command, signature, failure, and edit violations', () => {
  const nestedExtra = fixture('runtime-capabilities.json');
  nestedExtra.supportedKinds[0].unexpected = true;
  assertSchemaRejects('runtime-capabilities.schema.json', nestedExtra, 'nested additional property');

  const invalidVersion = fixture('runtime-capabilities.json');
  invalidVersion.package.version = 'release-1';
  assertSchemaRejects('runtime-capabilities.schema.json', invalidVersion, 'version pattern');

  const duplicateMediaType = fixture('runtime-capabilities.json');
  duplicateMediaType.supportedKinds[0].mediaTypes.push(duplicateMediaType.supportedKinds[0].mediaTypes[0]);
  assertSchemaRejects('runtime-capabilities.schema.json', duplicateMediaType, 'unique media types');

  const missingIdentifyCommand = fixture('runtime-capabilities.json');
  missingIdentifyCommand.commands = missingIdentifyCommand.commands.filter((command) => command.name !== 'identify');
  assertSchemaRejects('runtime-capabilities.schema.json', missingIdentifyCommand, 'required identify command');

  const missingCapabilitiesCommand = fixture('runtime-capabilities.json');
  missingCapabilitiesCommand.commands = missingCapabilitiesCommand.commands.filter((command) => command.name !== 'capabilities');
  assertSchemaRejects('runtime-capabilities.schema.json', missingCapabilitiesCommand, 'required capabilities command');

  const emptySupportedSignature = fixture('supported-identify.json');
  emptySupportedSignature.file.signature.evidence = [];
  assertSchemaRejects('runtime-evidence-envelope.schema.json', emptySupportedSignature, 'supported signature evidence');

  const missingFailureStage = fixture('failed-identify.json');
  delete missingFailureStage.failureStage;
  assertSchemaRejects('runtime-evidence-envelope.schema.json', missingFailureStage, 'failed stage');

  const missingFailureError = fixture('failed-identify.json');
  missingFailureError.errors = [];
  assertSchemaRejects('runtime-evidence-envelope.schema.json', missingFailureError, 'failed errors');

  const supportedWithoutSource = fixture('supported-identify.json');
  supportedWithoutSource.source = null;
  assertSchemaRejects('runtime-evidence-envelope.schema.json', supportedWithoutSource, 'supported source');

  const unsupportedWithoutSource = fixture('unsupported-identify.json');
  unsupportedWithoutSource.source = null;
  assertSchemaRejects('runtime-evidence-envelope.schema.json', unsupportedWithoutSource, 'unsupported source');

  const laterFailureWithoutSource = fixture('failed-identify.json');
  laterFailureWithoutSource.source = null;
  assertSchemaRejects('runtime-evidence-envelope.schema.json', laterFailureWithoutSource, 'later failure source');

  const sourceReadWithInventedSource = fixture('failed-source-read-identify.json');
  sourceReadWithInventedSource.source = fixture('failed-identify.json').source;
  assertSchemaRejects('runtime-evidence-envelope.schema.json', sourceReadWithInventedSource, 'source-read invented source');

  const sourceReadWithCheckedSignature = fixture('failed-source-read-identify.json');
  sourceReadWithCheckedSignature.file.signature.status = 'unknown';
  assertSchemaRejects('runtime-evidence-envelope.schema.json', sourceReadWithCheckedSignature, 'source-read signature');

  const nullAppliedPayload = fixture('edit-report.json');
  nullAppliedPayload.operations[0].appliedPayload = null;
  assertSchemaRejects('edit-report.schema.json', nullAppliedPayload, 'applied payload');

  const rejectedWithoutError = fixture('edit-report.json');
  rejectedWithoutError.operations[2].errors = [];
  assertSchemaRejects('edit-report.schema.json', rejectedWithoutError, 'rejected errors');
});

test('capability descriptor declares the non-mutating all-outcome identify probe', () => {
  validateCapabilities(fixture('runtime-capabilities.json'));
});

test('evidence fixtures distinguish supported, unsupported, and failed probes', () => {
  for (const name of [
    'supported-identify.json',
    'unsupported-identify.json',
    'failed-identify.json',
    'failed-source-read-identify.json',
  ]) {
    validateEvidence(fixture(name));
  }
});

test('failed evidence cannot claim a matched file kind', () => {
  const envelope = fixture('failed-identify.json');
  envelope.file = {
    fileKind: 'docx',
    mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    signature: { status: 'matched', kind: 'ooxml-content-type', evidence: ['wordprocessingml.document.main+xml'] },
  };

  assert.throws(() => validateEvidence(envelope), /failed file evidence/);
});

test('supported evidence requires non-empty runtime signature evidence', () => {
  const envelope = fixture('supported-identify.json');
  envelope.file.signature.evidence = [];

  assert.throws(() => validateEvidence(envelope), /signature evidence/);
});

test('derived evidence identities cannot masquerade as native ids', () => {
  const envelope = fixture('supported-identify.json');
  const derived = envelope.objects.find((object) => object.identity.kind === 'derived');
  derived.identity.nativeId = 'invented';
  assert.throws(() => validateEvidence(envelope), /must not fabricate nativeId/);
});

test('non-root evidence objects require a resolvable parent', () => {
  const envelope = fixture('supported-identify.json');
  envelope.objects[1].parentObjectId = 'missing-parent';
  assert.throws(() => validateEvidence(envelope), /parent must exist/);
});

test('evidence containment rejects parent cycles', () => {
  const envelope = fixture('supported-identify.json');
  envelope.objects[1].root = false;
  envelope.objects[1].parentObjectId = envelope.objects[2].objectId;
  envelope.objects[2].parentObjectId = envelope.objects[1].objectId;

  assert.throws(() => validateEvidence(envelope), /containment cycle/);
});

test('canonical evidence payloads reject language-dependent decimal numbers', () => {
  const envelope = fixture('supported-identify.json');
  envelope.payload = { value: 1.5 };

  assert.throws(() => validateEvidence(envelope), /integer numbers only/);
});

test('canonical JSON matches the shared cross-language adversarial vectors', () => {
  for (const vector of fixture('canonical-json-vectors.json').vectors) {
    assert.deepEqual(canonicalBytes(vector.value), Buffer.from(vector.canonical, 'utf8'), vector.name);
  }
});

test('canonical JSON rejects lossy numeric lexemes before JSON.parse', () => {
  for (const vector of fixture('canonical-json-negative-vectors.json').vectors) {
    assert.throws(() => canonicalBytesFromRaw(vector.json), /lexical integer|safe integer/, vector.name);
  }
});

test('the shared .NET contract project is explicitly discoverable as a test project', () => {
  const project = fs.readFileSync(path.join(here, '../../packages/_shared-dotnet.tests/runtime-contracts.tests.csproj'), 'utf8');
  assert.match(project, /<IsTestProject>true<\/IsTestProject>/);
});

test('edit report preserves one ordered result per request and derives its summary', () => {
  validateEditReport(fixture('edit-report.json'), fixture('edit-request.json'));
});

test('edit report rejects omitted applied payloads and non-derived summaries', () => {
  const request = fixture('edit-request.json');
  const missingPayload = fixture('edit-report.json');
  delete missingPayload.operations[0].appliedPayload;
  assert.throws(() => validateEditReport(missingPayload, request), /complete applied payload/);

  const wrongSummary = fixture('edit-report.json');
  wrongSummary.summary.applied = 2;
  assert.throws(() => validateEditReport(wrongSummary, request));
});

test('edit report cannot rewrite a request and attest to its own rewrite', () => {
  const report = fixture('edit-report.json');
  const request = fixture('edit-request.json');
  report.operations[0].requestedPayload.value = '43';
  const selfAuthoredRequest = { operations: report.operations.map((entry) => entry.requestedPayload) };
  const bytes = canonicalBytes(selfAuthoredRequest);
  const digest = createHash('sha256').update(bytes).digest('hex');
  report.requestArtifact.sizeBytes = bytes.byteLength;
  report.requestArtifact.sha256 = digest;
  report.requestArtifact.artifactId = `sha256:${digest}`;

  assert.throws(() => validateEditReport(report, request), /requestArtifact|authoritative request/);
});
