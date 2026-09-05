import test from 'node:test';
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {once} from 'node:events';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

test('stdio cancellation reaches an executing child and transport remains responsive', async () => {
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'tiwater-stdio-'));
  const marker=path.join(dir,'pid');
  const childCode=`require('fs').writeFileSync(${JSON.stringify(marker)},String(process.pid));setInterval(()=>{},1000)`;
  const script=`import {McpStdioServer} from ${JSON.stringify(new URL('./mcp-stdio.mjs',import.meta.url).href)};
import {runCandidateChain,commandCandidate} from ${JSON.stringify(new URL('./tool-runtime.mjs',import.meta.url).href)};
new McpStdioServer({name:'regression',version:'1',tools:[],callTool:async()=>runCandidateChain([commandCandidate(process.execPath)],['-e',${JSON.stringify(childCode)}])}).start();`;
  const server=spawn(process.execPath,['--input-type=module','-e',script],{stdio:['pipe','pipe','pipe']});
  let output='',stderr='';
  server.stdout.on('data',chunk=>output+=chunk);
  server.stderr.on('data',chunk=>stderr+=chunk);
  const close=once(server,'close');
  const send=message=>server.stdin.write(JSON.stringify({jsonrpc:'2.0',...message})+'\n');
  const waitFor=async predicate=>{
    const deadline=Date.now()+3000;
    while(!predicate()) {if(Date.now()>deadline) assert.fail(`Timed out: ${output} ${stderr}`);await new Promise(r=>setTimeout(r,10));}
  };
  try {
    send({id:1,method:'tools/call',params:{name:'wait'}});
    await waitFor(()=>fs.existsSync(marker));
    const pid=Number(fs.readFileSync(marker,'utf8'));
    send({method:'notifications/cancelled',params:{requestId:1}});
    send({id:2,method:'ping'});
    await waitFor(()=>output.includes('"id":2'));
    await waitFor(()=>{try {process.kill(pid,0);return false;} catch(error) {return error.code==='ESRCH';}});
    server.stdin.end();
    const [code]=await close;
    assert.equal(code,0,stderr);
    assert.ok(!output.includes('"stdout"'));
  } finally {server.kill('SIGKILL');fs.rmSync(dir,{recursive:true,force:true});}
});
