import { readFile } from 'node:fs/promises';
import path from 'node:path';

export async function resolveFileBackedChanges(args) {
  const hasInline = Array.isArray(args?.changes);
  const hasFile = typeof args?.changesInput === 'string' && args.changesInput.trim().length > 0;
  if (!hasInline && !hasFile) throw new Error('provide-changes-or-changesInput');
  if (hasInline && !hasFile) return args;

  let changes;
  try {
    changes = JSON.parse(await readFile(path.resolve(args.changesInput), 'utf8'));
  } catch (error) {
    throw new Error(`changesInput-invalid: ${error.message}`);
  }
  if (!Array.isArray(changes) || changes.length === 0) {
    throw new Error('changesInput-must-contain-non-empty-json-array');
  }
  const resolved = { ...args, changes: hasInline ? [...args.changes, ...changes] : changes };
  delete resolved.changesInput;
  return resolved;
}
