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
const project = path.join(repo, 'packages/docx-cli/docx.csproj');

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

test('capabilities --json emits a schema-valid non-mutating descriptor', () => {
  const result = runCli(['capabilities', '--json']);
  assert.equal(result.status, 0, result.stderr);
  const descriptor = JSON.parse(result.stdout);

  assertSchemaValid('runtime-capabilities.schema.json', descriptor);
  assert.deepEqual(descriptor.descriptorCommand, { command: 'capabilities', arguments: ['--json'], mutates: false });
  assert.deepEqual(descriptor.identifyProbe.outcomes, ['supported', 'unsupported', 'failed']);
  assert.ok(descriptor.commands.some(command => command.name === 'extract-evidence'));
});

test('identify missing source returns typed JSON failure and nonzero exit', () => {
  const missing = path.join(repo, `missing-${process.pid}.docx`);
  const result = runCli(['identify', missing, '--json']);
  assert.equal(result.status, 1, result.stderr);
  const evidence = JSON.parse(result.stdout);

  assertSchemaValid('runtime-evidence-envelope.schema.json', evidence);
  assert.equal(evidence.status, 'failed');
  assert.equal(evidence.failureStage, 'source-read');
  assert.equal(evidence.source, null);
  assert.equal(evidence.file.signature.status, 'not-checked');
  assert.ok(evidence.errors.length > 0);
});

test('identify fake docx returns unsupported JSON with a successful exit', () => {
  const fake = path.join(os.tmpdir(), `fake-${process.pid}-${Date.now()}.docx`);
  fs.writeFileSync(fake, 'not a zip package');
  try {
    const result = runCli(['identify', fake, '--json']);
    assert.equal(result.status, 0, result.stderr);
    const evidence = JSON.parse(result.stdout);
    assertSchemaValid('runtime-evidence-envelope.schema.json', evidence);
    assert.equal(evidence.status, 'unsupported');
    assert.equal(evidence.source.path, path.resolve(fake));
    assert.equal(evidence.file.fileKind, null);
    assert.deepEqual(evidence.errors, []);
  } finally {
    fs.rmSync(fake, { force: true });
  }
});

test('extract-evidence retains unsupported sources without parsing them', () => {
  const fake = path.join(os.tmpdir(), `fake-extract-${process.pid}-${Date.now()}.docx`);
  fs.writeFileSync(fake, 'not a zip package');
  try {
    const result = runCli(['extract-evidence', fake, '--json']);
    assert.equal(result.status, 0, result.stderr);
    const evidence = JSON.parse(result.stdout);
    assertSchemaValid('runtime-evidence-envelope.schema.json', evidence);
    assert.equal(evidence.probe, 'extract-evidence');
    assert.equal(evidence.status, 'unsupported');
    assert.deepEqual(evidence.objects, []);
  } finally {
    fs.rmSync(fake, { force: true });
  }
});
