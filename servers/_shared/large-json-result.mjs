import { createHash } from 'node:crypto';
import { createReadStream } from 'node:fs';
import { mkdir, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';

import { requireString } from './tool-runtime.mjs';

export const returnedContentBudgetBytes = 4_500;

export function resultChannels(args) {
  const returnContent = args.returnContent === true;
  const output = args.output === undefined ? null : path.resolve(requireString(args.output, 'output'));
  if (!returnContent && output === null) {
    throw Object.assign(new Error('returnContent-must-be-true-when-output-is-not-provided'), { code: -32602 });
  }
  return { returnContent, output };
}

export async function deliverLargeJsonResult({ tool, args, runtime, payload, sourcePaths, summary }) {
  const channels = resultChannels(args);
  const contentBytes = Buffer.byteLength(JSON.stringify(payload), 'utf8');
  const contentReturned = channels.returnContent && contentBytes <= returnedContentBudgetBytes;
  if (channels.returnContent && !contentReturned && channels.output === null) {
    throw Object.assign(
      new Error(`${tool}-content-is-${contentBytes}-bytes-provide-output-for-the-complete-result`),
      { code: -32602 },
    );
  }
  return {
    tool,
    runtime,
    sources: await Promise.all(sourcePaths.map(fileArtifact)),
    returnContent: channels.returnContent,
    artifact: channels.output === null ? null : await writeJsonArtifact(channels.output, payload),
    receipt: {
      contentBytes,
      contentReturned,
      contentWritten: channels.output !== null,
    },
    ...(summary === undefined ? {} : { summary }),
    ...(contentReturned ? { content: payload } : {}),
  };
}

export async function writeJsonArtifact(output, payload) {
  const fullPath = path.resolve(output);
  const bytes = Buffer.from(`${JSON.stringify(payload, null, 2)}\n`, 'utf8');
  await mkdir(path.dirname(fullPath), { recursive: true });
  await writeFile(fullPath, bytes, { flag: 'wx' });
  return {
    path: fullPath,
    sha256: createHash('sha256').update(bytes).digest('hex'),
    bytes: bytes.length,
  };
}

export async function fileArtifact(filePath) {
  const fullPath = path.resolve(filePath);
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(fullPath)) hash.update(chunk);
  const file = await stat(fullPath);
  return { path: fullPath, sha256: hash.digest('hex'), bytes: file.size };
}
