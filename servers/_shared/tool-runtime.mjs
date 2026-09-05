import fs from 'node:fs/promises';
import { constants as fsConstants } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { AsyncLocalStorage } from 'node:async_hooks';

const invocation = new AsyncLocalStorage();
const maxCommands = positiveInteger(process.env.TIWATER_MCP_MAX_COMMANDS ?? 4, 'TIWATER_MCP_MAX_COMMANDS');
const maxQueued = positiveInteger(process.env.TIWATER_MCP_MAX_QUEUED ?? 64, 'TIWATER_MCP_MAX_QUEUED');
let activeCommands = 0;
const commandQueue = [];

export function withCommandContext(context, fn) { return invocation.run(context, fn); }

function positiveInteger(value, name) {
  const result = Number(value);
  if (!Number.isSafeInteger(result) || result < 1 || result > 2_147_483_647) throw new Error(`${name} must be a positive bounded integer`);
  return result;
}

function commandError(code, message, executionStarted = false) {
  return Object.assign(new Error(message), { code, executionStarted });
}

function acquireCommand(signal) {
  if (signal.aborted) return Promise.reject(signal.reason);
  if (activeCommands < maxCommands) { activeCommands += 1; return Promise.resolve(releaseCommand); }
  if (commandQueue.length >= maxQueued) return Promise.reject(commandError('EBUSY', 'Command queue is full; execution did not start'));
  return new Promise((resolve, reject) => {
    const entry = { resolve, signal, abort: () => {
      const index = commandQueue.indexOf(entry);
      if (index >= 0) commandQueue.splice(index, 1);
      reject(signal.reason);
    } };
    commandQueue.push(entry);
    signal.addEventListener('abort', entry.abort, { once: true });
  });
}

function releaseCommand() {
  const next = commandQueue.shift();
  if (next) {
    next.signal.removeEventListener('abort', next.abort);
    next.resolve(releaseCommand);
  } else activeCommands -= 1;
}

const sharedDir = path.dirname(fileURLToPath(import.meta.url));
export const repoRoot = path.resolve(sharedDir, '..', '..');

export function createToolResult(payload, { isError = false } = {}) {
  return {
    ...(isError ? { isError: true } : {}),
    structuredContent: payload,
    content: isError
      ? [{ type: 'text', text: JSON.stringify(payload) }]
      : [],
  };
}

export function resolveRepoPath(...segments) {
  return path.join(repoRoot, ...segments);
}

export function commandCandidate(command, argsPrefix = [], options = {}) {
  return { command, argsPrefix, ...options };
}

export async function runCandidateChain(candidates, args, options = {}) {
  options = { ...invocation.getStore(), ...options };
  const timeoutMs = positiveInteger(options.timeoutMs ?? process.env.TIWATER_MCP_COMMAND_TIMEOUT_MS ?? 1_800_000, 'timeoutMs');
  const controller = new AbortController();
  const cancel = () => controller.abort(commandError('ABORT_ERR', 'Command cancelled; reconcile any started side effect before retrying'));
  if (options.signal?.aborted) cancel();
  else options.signal?.addEventListener('abort', cancel, { once: true });
  const deadline = setTimeout(() => controller.abort(commandError('ETIMEDOUT', 'Command deadline exceeded; reconcile any started side effect before retrying')), timeoutMs);
  try {
    return await runCandidates(candidates, args, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(deadline);
    options.signal?.removeEventListener('abort', cancel);
  }
}

async function runCandidates(candidates, args, options) {
  const errors = [];
  for (const candidate of candidates) {
    try {
      return await runCommand(candidate, args, options);
    } catch (error) {
      if (error?.code === 'ENOENT') {
        errors.push(`${candidate.command}: not found`);
        continue;
      }
      throw error;
    }
  }
  throw new Error(`No runnable command candidate succeeded. ${errors.join('; ')}`);
}

export async function runJsonCandidateChain(candidates, args, options = {}) {
  const result = await runCandidateChain(candidates, args, options);
  const text = result.stdout.trim();
  if (!text) return { ...result, json: null };
  try {
    return { ...result, json: JSON.parse(text) };
  } catch (error) {
    throw new Error(`Expected JSON output but received: ${text.slice(0, 300)}${text.length > 300 ? '…' : ''}`);
  }
}

export async function withTempJsonFile(data, fn) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'tiwater-mcp-'));
  const filePath = path.join(dir, 'payload.json');
  await fs.writeFile(filePath, JSON.stringify(data, null, 2), 'utf8');
  try {
    return await fn(filePath);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
}

export async function maybeReadJson(filePath) {
  const text = await fs.readFile(filePath, 'utf8');
  return JSON.parse(text);
}

export function requireString(value, label) {
  if (typeof value !== 'string' || value.trim() === '') {
    throw Object.assign(new Error(`${label} must be a non-empty string`), { code: -32602 });
  }
  return value;
}

async function runCommand(candidate, args, options) {
  const release = await acquireCommand(options.signal);
  try {
  options.signal.throwIfAborted();
  const env = await withDotnetRoot({ ...process.env, ...(candidate.env || {}), ...(options.env || {}) });
  options.signal.throwIfAborted();
  const cwd = candidate.cwd || options.cwd || repoRoot;
  const commandArgs = [...(candidate.argsPrefix || []), ...args];
  const maxOutputBytes = positiveInteger(options.maxOutputBytes ?? process.env.TIWATER_MCP_MAX_OUTPUT_BYTES ?? 67_108_864, 'maxOutputBytes');
  const killGraceMs = positiveInteger(options.killGraceMs ?? 1000, 'killGraceMs');

  return await new Promise((resolve, reject) => {
    const grouped = process.platform !== 'win32';
    const child = spawn(candidate.command, commandArgs, { cwd, env, detached: grouped, stdio: ['ignore', 'pipe', 'pipe'] });
    const stdoutChunks = [], stderrChunks = [];
    let outputBytes = 0, failure = null, killTimer = null;
    const kill = signal => {
      try { if (grouped && child.pid) process.kill(-child.pid, signal); else child.kill(signal); }
      catch (error) { if (error.code !== 'ESRCH') failure ||= error; }
    };
    const stop = error => {
      if (failure) return;
      failure = commandError(error.code || 'ABORT_ERR', error.message, Boolean(child.pid));
      kill('SIGTERM');
      killTimer = setTimeout(() => kill('SIGKILL'), killGraceMs);
    };
    const abort = () => stop(options.signal.reason);
    options.signal.addEventListener('abort', abort, { once: true });
    const collect = (chunks, chunk) => {
      if (failure) return;
      outputBytes += chunk.length;
      if (outputBytes > maxOutputBytes) { stop(commandError('ENOBUFS', 'Command output limit exceeded; output is incomplete')); return; }
      chunks.push(chunk);
    };
    child.stdout.on('data', chunk => collect(stdoutChunks, chunk));
    child.stderr.on('data', chunk => collect(stderrChunks, chunk));
    child.on('error', error => { failure ||= error; });
    child.on('close', code => {
      clearTimeout(killTimer);
      options.signal.removeEventListener('abort', abort);
      // A descendant may outlive the leader after closing inherited pipes.
      if (failure) { kill('SIGKILL'); reject(failure); return; }
      const stdout = Buffer.concat(stdoutChunks).toString('utf8');
      const stderr = Buffer.concat(stderrChunks).toString('utf8');
      const allowedExitCodes = options.allowedExitCodes ?? [0];
      if (allowedExitCodes.includes(code)) {
        resolve({ code, stdout, stderr, command: candidate.command, args: commandArgs, cwd });
        return;
      }
      reject(new Error(`${candidate.command} ${commandArgs.join(' ')} failed with exit code ${code}\n${stderr || stdout}`));
    });
  });
  } finally { release(); }
}

async function withDotnetRoot(env) {
  const architectureVariable = process.arch === 'arm64'
    ? 'DOTNET_ROOT_ARM64'
    : process.arch === 'x64'
      ? 'DOTNET_ROOT_X64'
      : null;
  if (env.DOTNET_ROOT || (architectureVariable && env[architectureVariable])) return env;

  const dotnet = await findOnPath(process.platform === 'win32' ? 'dotnet.exe' : 'dotnet', env.PATH);
  if (!dotnet) return env;

  const root = path.dirname(dotnet);
  return {
    ...env,
    DOTNET_ROOT: root,
    ...(architectureVariable ? { [architectureVariable]: root } : {}),
  };
}

async function findOnPath(command, pathValue) {
  for (const directory of String(pathValue ?? '').split(path.delimiter).filter(Boolean)) {
    const candidate = path.join(directory, command);
    try {
      await fs.access(candidate, fsConstants.X_OK);
      return candidate;
    } catch {
      // Continue to the next PATH entry.
    }
  }
  return null;
}
