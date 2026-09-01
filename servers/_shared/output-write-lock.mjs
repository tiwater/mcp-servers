import path from 'node:path';

const pendingWrites = new Map();

export async function withOutputWriteLock(output, write) {
  const target = path.resolve(output);
  const previous = pendingWrites.get(target) ?? Promise.resolve();
  const current = previous.catch(() => undefined).then(write);
  pendingWrites.set(target, current);
  try {
    return await current;
  } finally {
    if (pendingWrites.get(target) === current) pendingWrites.delete(target);
  }
}
