import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import {
  commandCandidate,
  redactCommandArgs,
  runCandidateChain,
} from './tool-runtime.mjs';

async function withFixture(contents, fn) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'tiwater-runtime-test-'));
  const script = path.join(directory, 'candidate.mjs');
  await fs.writeFile(script, contents, 'utf8');
  try {
    return await fn(script);
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
}

test('skips an installed candidate whose capability identity does not match', async () => {
  await withFixture(`
    const command = process.argv[2];
    if (command === 'capabilities') {
      console.log(JSON.stringify({ descriptorType: 'runtime-capabilities', runtime: { name: process.argv[4] } }));
    } else {
      console.log(process.argv[3]);
    }
  `, async script => {
    const candidates = [
      commandCandidate(process.execPath, [script], { expectedRuntimeName: 'tiwater-docx', capabilityArgs: ['capabilities', '--json', 'wrong-runtime'] }),
      commandCandidate(process.execPath, [script], { expectedRuntimeName: 'tiwater-docx', capabilityArgs: ['capabilities', '--json', 'tiwater-docx'] }),
    ];

    const result = await runCandidateChain(candidates, ['run', 'selected-fallback']);

    assert.equal(result.stdout.trim(), 'selected-fallback');
    assert.equal(result.capabilities.runtime.name, 'tiwater-docx');
  });
});

test('redacts secret option values from reported command arguments', () => {
  assert.deepEqual(
    redactCommandArgs(
      ['extract-tables', 'input.pdf', '--api-key', 'secret-value', '--llm-model', 'model'],
      ['--api-key'],
    ),
    ['extract-tables', 'input.pdf', '--api-key', '[REDACTED]', '--llm-model', 'model'],
  );
});

test('does not expose secret option values when a command fails', async () => {
  await withFixture('process.exit(2);', async script => {
    const candidate = commandCandidate(process.execPath, [script], {
      secretOptions: ['--api-key'],
    });

    await assert.rejects(
      runCandidateChain([candidate], ['--api-key', 'secret-value']),
      error => !error.message.includes('secret-value') && error.message.includes('[REDACTED]'),
    );
  });
});
