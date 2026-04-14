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
      if (code === 0) {
        resolve({ code, stdout, stderr, command: candidate.command, args: commandArgs });
        return;
      }
      reject(new Error(`${candidate.command} ${commandArgs.join(' ')} failed with exit code ${code}\n${stderr || stdout}`));
    });
  });
}
