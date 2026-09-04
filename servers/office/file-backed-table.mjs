import { readFile } from 'node:fs/promises';
import path from 'node:path';

const tableFields = ['table', 'existingRows', 'columns', 'rows'];

function hasInlineTable(args) {
  return tableFields.every((field) => args?.[field] !== undefined);
}

export async function resolveFileBackedTable(args) {
  const hasInline = hasInlineTable(args);
  const hasFile = typeof args?.tableInput === 'string' && args.tableInput.trim().length > 0;
  if (!hasInline && !hasFile) throw new Error('provide-inline-table-or-tableInput');
  if (hasInline && hasFile) throw new Error('provide-only-one-table-input');
  if (hasInline) return args;

  let tableRequest;
  try {
    tableRequest = JSON.parse(await readFile(path.resolve(args.tableInput), 'utf8'));
  } catch (error) {
    throw new Error(`tableInput-invalid: ${error.message}`);
  }
  if (!tableRequest || typeof tableRequest !== 'object' || Array.isArray(tableRequest)
      || !tableFields.every((field) => tableRequest[field] !== undefined)
      || Object.keys(tableRequest).some((field) => !tableFields.includes(field))) {
    throw new Error('tableInput-must-contain-table-existingRows-columns-and-rows');
  }
  const resolved = { ...args, ...tableRequest };
  delete resolved.tableInput;
  return resolved;
}
