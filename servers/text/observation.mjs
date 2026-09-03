import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { fileArtifact } from '../_shared/large-json-result.mjs';

const supportedExtensions = new Set(['.txt', '.text', '.log', '.csv', '.tsv', '.md', '.markdown']);
const openingLineLimit = 8;
const openingTextLimit = 160;
const linePageLimit = 200;

export async function inspectText(inputValue) {
  const observation = await observeText(inputValue);
  const source = await fileArtifact(observation.input);
  const openingLines = observation.lines.slice(0, openingLineLimit).map(line => ({
    identity: { sourceSha256: source.sha256, index: line.index },
    textPreview: preview(line.text, openingTextLimit),
    textLength: [...line.text].length,
    terminator: line.terminator,
  }));
  const identity = {
    source,
    extension: observation.extension,
    decoding: observation.decoding,
    lineCount: observation.lines.length,
    openingLines,
  };
  return {
    input: observation.input,
    identity,
    payload: { schema: 'tiwater.text-inspection/v1', ...identity },
  };
}

export async function readTextLines(inputValue, requestedOffset) {
  if (!Number.isSafeInteger(requestedOffset) || requestedOffset < 0) {
    throw Object.assign(new Error('offset must be a non-negative safe integer'), { code: -32602 });
  }
  const observation = await observeText(inputValue);
  const source = await fileArtifact(observation.input);
  const offset = Math.min(requestedOffset, observation.lines.length);
  const selected = observation.lines.slice(offset, offset + linePageLimit);
  const nextOffset = offset + selected.length < observation.lines.length
    ? offset + selected.length
    : null;
  const receipt = {
    schema: 'tiwater.text-line-page-receipt/v1',
    totalLineCount: observation.lines.length,
    returnedLineCount: selected.length,
    remaining: observation.lines.length - offset - selected.length,
    nextOffset,
  };
  return {
    input: observation.input,
    receipt,
    payload: {
      schema: 'tiwater.text-line-page/v1',
      source,
      extension: observation.extension,
      decoding: observation.decoding,
      receipt,
      lines: selected.map(line => ({
        identity: { sourceSha256: source.sha256, index: line.index },
        text: line.text,
        terminator: line.terminator,
      })),
    },
  };
}

export async function observeText(inputValue) {
  if (typeof inputValue !== 'string' || inputValue.trim() === '') {
    throw Object.assign(new Error('input must be a non-empty string'), { code: -32602 });
  }
  const input = path.resolve(inputValue);
  const extension = path.extname(input).toLowerCase();
  if (!supportedExtensions.has(extension)) {
    throw Object.assign(new Error(`unsupported-plain-text-extension:${extension || '(none)'}`), { code: -32602 });
  }
  const bytes = await readFile(input);
  const decoded = decodeLosslessly(bytes);
  rejectBinaryControls(decoded.text);
  return { input, extension, decoding: decoded.decoding, lines: splitLines(decoded.text) };
}

function decodeLosslessly(bytes) {
  if (bytes.subarray(0, 3).equals(Buffer.from([0xef, 0xbb, 0xbf]))) {
    return decodeWithRoundTrip(bytes.subarray(3), 'utf-8', 'utf-8');
  }
  if (bytes.subarray(0, 2).equals(Buffer.from([0xff, 0xfe]))) {
    return decodeWithRoundTrip(bytes.subarray(2), 'utf-16le', 'utf-16le');
  }
  if (bytes.subarray(0, 2).equals(Buffer.from([0xfe, 0xff]))) {
    return decodeWithRoundTrip(bytes.subarray(2), 'utf-16be', 'utf-16be');
  }
  return decodeWithRoundTrip(bytes, 'utf-8', 'none');
}

function decodeWithRoundTrip(bytes, encoding, bom) {
  if ((encoding === 'utf-16le' || encoding === 'utf-16be') && bytes.length % 2 !== 0) {
    throw Object.assign(new Error(`invalid-${encoding}-byte-length`), { code: -32602 });
  }
  let text;
  try {
    text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
  } catch {
    throw Object.assign(new Error(`invalid-${encoding}-sequence`), { code: -32602 });
  }
  let encoded = encoding === 'utf-8' ? Buffer.from(text, 'utf8') : Buffer.from(text, 'utf16le');
  if (encoding === 'utf-16be') encoded = swapUtf16Bytes(encoded);
  if (!encoded.equals(bytes)) {
    throw Object.assign(new Error(`non-lossless-${encoding}-decode`), { code: -32602 });
  }
  return { text, decoding: { status: 'lossless', encoding, bom } };
}

function swapUtf16Bytes(bytes) {
  const swapped = Buffer.allocUnsafe(bytes.length);
  for (let index = 0; index < bytes.length; index += 2) {
    swapped[index] = bytes[index + 1];
    swapped[index + 1] = bytes[index];
  }
  return swapped;
}

function rejectBinaryControls(text) {
  const binary = [...text].some(character => {
    const code = character.codePointAt(0);
    return code <= 8 || code === 11 || code === 12 || (code >= 14 && code <= 31)
      || (code >= 127 && code <= 159);
  });
  if (binary) throw Object.assign(new Error('binary-control-content-is-not-plain-text'), { code: -32602 });
}

function splitLines(text) {
  if (text.length === 0) return [];
  const lines = [];
  let start = 0;
  while (start < text.length) {
    let end = start;
    while (end < text.length && text[end] !== '\r' && text[end] !== '\n') end++;
    let terminator = 'none';
    let next = end;
    if (end < text.length) {
      if (text[end] === '\r' && text[end + 1] === '\n') {
        terminator = 'crlf';
        next = end + 2;
      } else {
        terminator = text[end] === '\r' ? 'cr' : 'lf';
        next = end + 1;
      }
    }
    lines.push({ index: lines.length, text: text.slice(start, end), terminator });
    start = next;
  }
  return lines;
}

function preview(text, limit) {
  return [...text].slice(0, limit).join('');
}
