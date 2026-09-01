import assert from 'node:assert/strict';
import test from 'node:test';
import { withOutputWriteLock } from './output-write-lock.mjs';

function deferred() {
  let resolve;
  const promise = new Promise(done => { resolve = done; });
  return { promise, resolve };
}

test('serializes writes to the same normalized output path', async () => {
  const firstMayFinish = deferred();
  const events = [];
  const first = withOutputWriteLock('/tmp/report.docx', async () => {
    events.push('first-start');
    await firstMayFinish.promise;
    events.push('first-end');
  });
  const second = withOutputWriteLock('/tmp/./report.docx', async () => {
    events.push('second-start');
  });

  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(events, ['first-start']);
  firstMayFinish.resolve();
  await Promise.all([first, second]);
  assert.deepEqual(events, ['first-start', 'first-end', 'second-start']);
});

test('allows writes to different output paths to run concurrently', async () => {
  const mayFinish = deferred();
  const events = [];
  const first = withOutputWriteLock('/tmp/first.docx', async () => {
    events.push('first-start');
    await mayFinish.promise;
  });
  const second = withOutputWriteLock('/tmp/second.docx', async () => {
    events.push('second-start');
    await mayFinish.promise;
  });

  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(events.sort(), ['first-start', 'second-start']);
  mayFinish.resolve();
  await Promise.all([first, second]);
});

test('continues the queue after a failed write', async () => {
  const events = [];
  const first = withOutputWriteLock('/tmp/retry.docx', async () => {
    events.push('first');
    throw new Error('write failed');
  });
  const second = withOutputWriteLock('/tmp/retry.docx', async () => {
    events.push('second');
    return 'ok';
  });

  await assert.rejects(first, /write failed/);
  assert.equal(await second, 'ok');
  assert.deepEqual(events, ['first', 'second']);
});
