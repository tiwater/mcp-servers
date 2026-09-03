import { spawn } from 'node:child_process';

export function createToolResult(payload, { isError = false } = {}) {
  return {
    ...(isError ? { isError: true } : {}),
    structuredContent: payload,
    content: isError ? [{ type: 'text', text: JSON.stringify(payload) }] : [],
  };
}

export function commandCandidate(command, argsPrefix = [], options = {}) {
  return { command, argsPrefix, ...options };
}

export async function runJsonCandidateChain(candidates, args, options = {}) {
  const errors = [];
  for (const candidate of candidates) {
    try {
      const result = await runCommand(candidate, args, options);
      const text = result.stdout.trim();
      if (!text) return { ...result, json: null };
      try {
        return { ...result, json: JSON.parse(text) };
      } catch {
        throw new Error(`Expected JSON output but received: ${text.slice(0, 300)}${text.length > 300 ? '…' : ''}`);
      }
    } catch (error) {
      if (error?.code !== 'ENOENT') throw error;
      errors.push(`${candidate.command}: not found`);
    }
  }
  throw new Error(`No runnable command candidate succeeded. ${errors.join('; ')}`);
}

export function requireString(value, label) {
  if (typeof value !== 'string' || value.trim() === '') {
    throw Object.assign(new Error(`${label} must be a non-empty string`), { code: -32602 });
  }
  return value;
}

function runCommand(candidate, args, options) {
  const commandArgs = [...(candidate.argsPrefix || []), ...args];
  return new Promise((resolve, reject) => {
    const child = spawn(candidate.command, commandArgs, {
      cwd: candidate.cwd || options.cwd || process.cwd(),
      env: { ...process.env, ...(candidate.env || {}), ...(options.env || {}) },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', chunk => { stdout += chunk.toString('utf8'); });
    child.stderr.on('data', chunk => { stderr += chunk.toString('utf8'); });
    child.on('error', reject);
    child.on('close', code => {
      if (code === 0) {
        resolve({ command: candidate.command, args: commandArgs, stdout, stderr });
      } else {
        const error = new Error(`${candidate.command} exited with code ${code}: ${stderr.trim()}`);
        error.code = code;
        reject(error);
      }
    });
  });
}
