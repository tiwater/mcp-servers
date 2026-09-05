import test from 'node:test';
import assert from 'node:assert/strict';
import { runCandidateChain, commandCandidate } from './tool-runtime.mjs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
const node = commandCandidate(process.execPath);
test('bounded commands preserve output and allowed nonzero exits', async () => {
  const result = await runCandidateChain([node], ['-e', 'process.stdout.write("中文");process.stderr.write("note");process.exitCode=3'], {allowedExitCodes:[3]});
  assert.equal(result.stdout, '中文');
  assert.equal(result.stderr, 'note');
  assert.equal(result.code, 3);
});
test('deadline terminates a command instead of accepting late success', async () => {
  await assert.rejects(runCandidateChain([node], ['-e', 'setTimeout(()=>{},400)'], {timeoutMs:80, killGraceMs:30}), {code:'ETIMEDOUT'});
});
test('pre-cancelled command never executes or falls back', async () => {
  const controller=new AbortController(); controller.abort();
  await assert.rejects(runCandidateChain([node,node], ['-e','process.stdout.write("executed")'], {signal:controller.signal}), {code:'ABORT_ERR'});
});
test('running command cancellation waits for process termination', async () => {
  const controller=new AbortController();
  const timer=setTimeout(()=>controller.abort(),80);
  try { await assert.rejects(runCandidateChain([node], ['-e','setTimeout(()=>{},400)'], {signal:controller.signal,killGraceMs:30}), {code:'ABORT_ERR'}); }
  finally { clearTimeout(timer); }
});
test('output limit fails without silently truncating successful JSON', async () => {
  await assert.rejects(runCandidateChain([node], ['-e','process.stdout.write("a".repeat(4096))'], {maxOutputBytes:1024}), {code:'ENOBUFS'});
});
test('only missing executable can select another candidate', async () => {
  const result=await runCandidateChain([commandCandidate('/missing-tiwater-regression-command'),node], ['-e','process.stdout.write("ok")']);
  assert.equal(result.stdout,'ok');
  await assert.rejects(runCandidateChain([node,node], ['-e','process.exit(7)']), /exit code 7/);
});
test('queued deadline cannot start a command after its caller has timed out', async () => {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'tiwater-queue-'));
  const marker=path.join(dir,'started');
  const busy=Array.from({length:4},()=>runCandidateChain([node],['-e','setTimeout(()=>{},250)']));
  try {
    await assert.rejects(runCandidateChain([node],['-e',`require('fs').writeFileSync(${JSON.stringify(marker)},'bad')`],{timeoutMs:30}),{code:'ETIMEDOUT',executionStarted:false});
    await Promise.all(busy);
    assert.equal(fs.existsSync(marker),false);
  } finally { await Promise.allSettled(busy); fs.rmSync(dir,{recursive:true,force:true}); }
});
test('timeout kills a descendant which ignores termination', {skip:process.platform==='win32'}, async () => {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'tiwater-child-'));
  const marker=path.join(dir,'pid');
  const script=`const cp=require('child_process'); const c=cp.spawn(process.execPath,['-e',"process.on('SIGTERM',()=>{});setInterval(()=>{},1000)"],{stdio:'ignore'});require('fs').writeFileSync(${JSON.stringify(marker)},String(c.pid));setInterval(()=>{},1000);`;
  try {
    await assert.rejects(runCandidateChain([node],['-e',script],{timeoutMs:250,killGraceMs:30}),{code:'ETIMEDOUT',executionStarted:true});
    const pid=Number(fs.readFileSync(marker,'utf8'));
    // SIGKILL delivery and orphan reaping can lag the leader's close event.
    const deadline=Date.now()+1000;
    while(Date.now()<deadline) {
      try {process.kill(pid,0);} catch(error) {assert.equal(error.code,'ESRCH');return;}
      await new Promise(resolve=>setTimeout(resolve,20));
    }
    assert.fail(`descendant ${pid} survived cancellation`);
  } finally {fs.rmSync(dir,{recursive:true,force:true});}
});
