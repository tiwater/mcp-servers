import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';

const sharedDir = path.dirname(fileURLToPath(import.meta.url));
export const repoRoot = path.resolve(sharedDir, '..', '..');

export function createToolResult(payload) {
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify(payload, null, 2),
      },
    ],
  };
}

export function resolveRepoPath(...segments) {
  return path.join(repoRoot, ...segments);
}

export function commandCandidate(command, argsPrefix = [], options = {}) {
  return { command, argsPrefix, ...options };
}

export async function runCandidateChain(candidates, args, options = {}) {
  const errors = [];
  for (const candidate of candidates) {
    try {
      const capabilities = await qualifyCandidate(candidate, options);
      const result = await runCommand(candidate, args, options);
      return capabilities ? { ...result, capabilities } : result;
    } catch (error) {
      if (error?.code === 'ENOENT' || error?.code === 'EUNQUALIFIED') {
        errors.push(`${candidate.command}: ${error.code === 'ENOENT' ? 'not found' : error.message}`);
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

export function redactCommandArgs(args, secretOptions = []) {
  const secrets = new Set(secretOptions);
  return args.map((value, index) => (
    index > 0 && secrets.has(args[index - 1]) ? '[REDACTED]' : value
  ));
}

async function qualifyCandidate(candidate, options) {
  if (!candidate.expectedRuntimeName) return null;

  let result;
  try {
    result = await runCommand(
      candidate,
      candidate.capabilityArgs || ['capabilities', '--json'],
      { ...options, acceptedExitCodes: [0] },
    );
  } catch (error) {
    if (error?.code === 'ENOENT') throw error;
    throw unqualified(`capability probe failed: ${error.message}`);
  }

  let descriptor;
  try {
    descriptor = JSON.parse(result.stdout);
  } catch {
    throw unqualified('capability probe did not return JSON');
  }

  if (descriptor?.descriptorType !== 'runtime-capabilities') {
    throw unqualified('capability descriptor type mismatch');
  }
  if (descriptor?.runtime?.name !== candidate.expectedRuntimeName) {
    throw unqualified(`runtime identity mismatch (expected ${candidate.expectedRuntimeName})`);
  }
  return descriptor;
}

function unqualified(message) {
  return Object.assign(new Error(message), { code: 'EUNQUALIFIED' });
}

async function runCommand(candidate, args, options) {
  const env = { ...process.env, ...(candidate.env || {}), ...(options.env || {}) };
  const cwd = candidate.cwd || options.cwd || repoRoot;
  const commandArgs = [...(candidate.argsPrefix || []), ...args];

  return await new Promise((resolve, reject) => {
    const child = spawn(candidate.command, commandArgs, { cwd, env, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';

    child.stdout.on('data', chunk => {
      stdout += chunk.toString();
    });

    child.stderr.on('data', chunk => {
      stderr += chunk.toString();
    });

    child.on('error', reject);
    child.on('close', code => {
      const acceptedExitCodes = options.acceptedExitCodes || [0];
      if (acceptedExitCodes.includes(code)) {
        resolve({ code, stdout, stderr, command: candidate.command, args: commandArgs });
        return;
      }
      const safeArgs = redactCommandArgs(commandArgs, candidate.secretOptions || []);
      const safeDetails = redactSecretsInText(stderr || stdout, commandArgs, candidate.secretOptions || []);
      reject(new Error(`${candidate.command} ${safeArgs.join(' ')} failed with exit code ${code}\n${safeDetails}`));
    });
  });
}

function redactSecretsInText(text, args, secretOptions) {
  let result = text;
  const options = new Set(secretOptions);
  for (let index = 1; index < args.length; index += 1) {
    if (options.has(args[index - 1]) && args[index]) {
      result = result.replaceAll(args[index], '[REDACTED]');
    }
  }
  return result;
}
