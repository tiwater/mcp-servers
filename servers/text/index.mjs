#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import * as z from 'zod/v4';
import { McpServer } from '@modelcontextprotocol/server';
import { serveStdio } from '@modelcontextprotocol/server/stdio';

import { createToolResult } from '../_shared/tool-runtime.mjs';
import { deliverLargeJsonResult } from '../_shared/large-json-result.mjs';
import { withOutputWriteLock } from '../_shared/output-write-lock.mjs';
import { evidenceRoleMetadata } from '../_shared/evidence-role.mjs';
import { inspectText, readTextLines } from './observation.mjs';

const packageMetadata = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
const contractManifest = JSON.parse(await readFile(
  new URL('./contracts/tiwater-text-provider-contract-manifest-v1.json', import.meta.url),
  'utf8',
));
if (contractManifest.schema !== 'tiwater.text-provider-contract-manifest/v1'
    || contractManifest.provider?.id !== packageMetadata.name
    || contractManifest.provider?.version !== packageMetadata.version) {
  throw new Error('Text MCP input contract manifest does not match the installed distribution');
}

const inputContracts = new Map(await Promise.all(contractManifest.tools.map(async entry => {
  const bytes = await readFile(new URL(`./contracts/${entry.name}.schema.json`, import.meta.url));
  const hash = createHash('sha256').update(bytes).digest('hex');
  if (hash !== entry.inputContract.sha256) {
    throw new Error(`Text MCP input contract hash mismatch: ${entry.name}`);
  }
  return [entry.name, z.fromJSONSchema(JSON.parse(bytes.toString('utf8')))];
})));

const openingLineLimit = 8;
const artifact = z.object({
  path: z.string(),
  sha256: z.string().regex(/^[0-9a-f]{64}$/),
  bytes: z.number().int().nonnegative(),
}).strict();
const runtimeIdentity = z.object({ command: z.literal('tiwater-text-mcp'), cwd: z.string() }).strict();
const decoding = z.object({
  status: z.literal('lossless'),
  encoding: z.enum(['utf-8', 'utf-16le', 'utf-16be']),
  bom: z.enum(['none', 'utf-8', 'utf-16le', 'utf-16be']),
}).strict();
const lineIdentity = z.object({
  sourceSha256: z.string().regex(/^[0-9a-f]{64}$/),
  index: z.number().int().nonnegative(),
}).strict();
const terminator = z.enum(['none', 'lf', 'crlf', 'cr']);
const openingLine = z.object({
  identity: lineIdentity,
  textPreview: z.string(),
  textLength: z.number().int().nonnegative(),
  terminator,
}).strict();
const inspectIdentity = z.object({
  source: artifact,
  extension: z.string(),
  decoding,
  lineCount: z.number().int().nonnegative(),
  openingLines: z.array(openingLine).max(openingLineLimit),
}).strict();
const linePageReceipt = z.object({
  schema: z.literal('tiwater.text-line-page-receipt/v1'),
  totalLineCount: z.number().int().nonnegative(),
  returnedLineCount: z.number().int().nonnegative(),
  remaining: z.number().int().nonnegative(),
  nextOffset: z.number().int().nonnegative().nullable(),
}).strict();
const textLine = z.object({ identity: lineIdentity, text: z.string(), terminator }).strict();
const inspectContent = z.object({
  schema: z.literal('tiwater.text-inspection/v1'),
  source: artifact,
  extension: z.string(),
  decoding,
  lineCount: z.number().int().nonnegative(),
  openingLines: z.array(openingLine).max(openingLineLimit),
}).strict();
const linePage = z.object({
  schema: z.literal('tiwater.text-line-page/v1'),
  source: artifact,
  extension: z.string(),
  decoding,
  receipt: linePageReceipt,
  lines: z.array(textLine).max(200),
}).strict();

function largeResultOutput(contentSchema) {
  return z.object({
    tool: z.string(),
    runtime: runtimeIdentity,
    sources: z.array(artifact).min(1).max(1),
    returnContent: z.boolean(),
    artifact: artifact.nullable(),
    receipt: z.object({
      contentBytes: z.number().int().nonnegative(),
      contentReturned: z.boolean(),
      contentWritten: z.boolean(),
    }).strict(),
    content: contentSchema.optional(),
  }).strict();
}

const definitions = [
  {
    name: 'text_inspect',
    evidenceRole: 'document-observation',
    description: 'Inspect one exact supported plain-text revision. Return its byte identity, lossless encoding and BOM facts, line count, and at most eight opening line identities while retaining the complete bounded inspection at output. It does not parse fields, records, key-value pairs, sections, or markup.',
    outputSchema: largeResultOutput(inspectContent).extend({ identity: inspectIdentity }).strict(),
    handler: textInspect,
  },
  {
    name: 'text_read_lines',
    description: 'Read one explicit zero-based line page from one exact supported plain-text revision. The receipt reports remaining lines and nextOffset; continue only when another line is needed. Set returnContent true to return the selected page when it fits the response limit. Provide output to retain the same complete page. These channels are independent and may be used together; at least one is required. Lines retain their exact decoded text and terminator; the provider does not interpret fields, records, key-value pairs, sections, Markdown, or business meaning.',
    outputSchema: largeResultOutput(linePage).extend({ summary: linePageReceipt }).strict(),
    handler: textReadLines,
  },
];

function buildServer() {
  const server = new McpServer(
    { name: 'tiwater-text', version: packageMetadata.version },
    { instructions: 'Observe only exact supported plain-text bytes and explicit zero-based line pages. A read-only output path is an immutable artifact identity: an identical request may replay it; every different request uses a different path. Callers own all interpretation and business meaning.' },
  );
  for (const definition of definitions) {
    const inputSchema = inputContracts.get(definition.name);
    if (!inputSchema) throw new Error(`Missing provider-owned Text MCP input contract: ${definition.name}`);
    server.registerTool(
      definition.name,
      {
        description: definition.description,
        inputSchema,
        outputSchema: definition.outputSchema,
        annotations: {
          readOnlyHint: true,
          idempotentHint: true,
          destructiveHint: false,
          openWorldHint: false,
        },
        ...(definition.evidenceRole ? { _meta: evidenceRoleMetadata(definition.evidenceRole) } : {}),
      },
      async args => {
        const payload = typeof args.output === 'string'
          ? await withOutputWriteLock(args.output, () => definition.handler(args))
          : await definition.handler(args);
        return createToolResult(payload);
      },
    );
  }
  return server;
}

async function textInspect(args) {
  const observation = await inspectText(args.input);
  const delivered = await deliverLargeJsonResult({
    tool: 'text_inspect',
    args,
    runtime: runtime(),
    payload: observation.payload,
    sourcePaths: [observation.input],
  });
  return { ...delivered, identity: observation.identity };
}

async function textReadLines(args) {
  const observation = await readTextLines(args.input, args.offset, args.limit);
  const delivered = await deliverLargeJsonResult({
    tool: 'text_read_lines',
    args,
    runtime: runtime(),
    payload: observation.payload,
    sourcePaths: [observation.input],
  });
  return { ...delivered, summary: observation.receipt };
}

function runtime() {
  return { command: 'tiwater-text-mcp', cwd: process.cwd() };
}

serveStdio(buildServer);
