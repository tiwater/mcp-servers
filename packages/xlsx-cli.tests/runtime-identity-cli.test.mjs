import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';
import Ajv2020 from 'ajv/dist/2020.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const project = path.join(repo, 'packages/xlsx-cli/xlsx.csproj');

function runCli(args) {
  return spawnSync('dotnet', ['run', '--project', project, '-c', 'Release', '--', ...args], {
    cwd: repo,
    encoding: 'utf8',
  });
}

function assertSchemaValid(schemaName, value) {
  const schema = JSON.parse(fs.readFileSync(path.join(repo, `contracts/runtime/${schemaName}`)));
  const validate = new Ajv2020({ allErrors: true, strict: false }).compile(schema);
  assert.equal(validate(value), true, JSON.stringify(validate.errors));
}

test('capabilities --json emits the schema-valid non-mutating XLS/XLSX descriptor', () => {
  const result = runCli(['capabilities', '--json']);
  assert.equal(result.status, 0, result.stderr);
  const descriptor = JSON.parse(result.stdout);

  assertSchemaValid('runtime-capabilities.schema.json', descriptor);
  assert.deepEqual(descriptor.descriptorCommand, { command: 'capabilities', arguments: ['--json'], mutates: false });
  assert.deepEqual(descriptor.identifyProbe, {
    command: 'identify',
    arguments: ['<input>', '--json'],
    mutates: false,
    outcomes: ['supported', 'unsupported', 'failed'],
  });
  assert.deepEqual(descriptor.supportedKinds.map(({ fileKind }) => fileKind), ['xlsx', 'xls']);
});

test('identify missing source emits a schema-valid source-read failure', () => {
  const missing = path.join(repo, `missing-${process.pid}.xlsx`);
  const result = runCli(['identify', missing, '--json']);
  assert.equal(result.status, 1, result.stderr);
  const evidence = JSON.parse(result.stdout);

  assertSchemaValid('runtime-evidence-envelope.schema.json', evidence);
  assert.equal(evidence.status, 'failed');
  assert.equal(evidence.failureStage, 'source-read');
  assert.equal(evidence.source, null);
  assert.equal(evidence.file.signature.status, 'not-checked');
  assert.deepEqual(evidence.file.signature.evidence, []);
});

test('identify fake extension emits schema-valid unsupported evidence', () => {
  const fake = path.join(os.tmpdir(), `fake-${process.pid}-${Date.now()}.xlsx`);
  fs.writeFileSync(fake, 'not a spreadsheet package');
  try {
    const result = runCli(['identify', fake, '--json']);
    assert.equal(result.status, 0, result.stderr);
    const evidence = JSON.parse(result.stdout);

    assertSchemaValid('runtime-evidence-envelope.schema.json', evidence);
    assert.equal(evidence.status, 'unsupported');
    assert.equal(evidence.source.path, path.resolve(fake));
    assert.equal(evidence.file.fileKind, null);
    assert.equal(evidence.file.mediaType, null);
    assert.deepEqual(evidence.errors, []);
  } finally {
    fs.rmSync(fake, { force: true });
  }
});
