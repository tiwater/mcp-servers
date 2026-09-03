#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { deliverLargeJsonResult } from '../_shared/large-json-result.mjs';
import { McpStdioServer } from '../_shared/mcp-stdio.mjs';
import {
  commandCandidate,
  createToolResult,
  requireString,
  runJsonCandidateChain,
} from '../_shared/tool-runtime.mjs';

const pdfRoot = path.dirname(fileURLToPath(import.meta.url));
const distributionRoot = path.resolve(pdfRoot, '..');
const packageJson = readJson(path.join(distributionRoot, 'package.json'));
const manifest = readJson(path.join(
  pdfRoot, 'contracts', 'tiwater-pdf-provider-contract-manifest-v1.json',
));
if (manifest.schema !== 'tiwater.pdf-provider-contract-manifest/v1'
  || manifest.provider?.id !== packageJson.name
  || manifest.provider?.version !== packageJson.version) {
  throw new Error('pdf-provider-contract-manifest-invalid');
}

const pdfCandidates = [commandCandidate('tiwater-pdf')];
const artifactSchema = {
  type: 'object',
  properties: {
    path: { type: 'string' },
    sha256: { type: 'string', pattern: '^[0-9a-f]{64}$' },
    bytes: { type: 'integer', minimum: 1 },
  },
  required: ['path', 'sha256', 'bytes'],
  additionalProperties: false,
};
const receiptSchema = {
  type: 'object',
  properties: {
    contentBytes: { type: 'integer', minimum: 0 },
    contentReturned: { type: 'boolean' },
    contentWritten: { type: 'boolean' },
  },
  required: ['contentBytes', 'contentReturned', 'contentWritten'],
  additionalProperties: false,
};
const largeResultOutputSchema = {
  type: 'object',
  properties: {
    tool: { type: 'string' },
    runtime: {
      type: 'object',
      properties: {
        command: { type: 'string' },
        arguments: { type: 'array', items: { type: 'string' } },
      },
      required: ['command', 'arguments'],
      additionalProperties: false,
    },
    sources: { type: 'array', minItems: 1, maxItems: 1, items: artifactSchema },
    returnContent: { type: 'boolean' },
    artifact: { anyOf: [artifactSchema, { type: 'null' }] },
    receipt: receiptSchema,
    content: {},
  },
  required: ['tool', 'runtime', 'sources', 'returnContent', 'artifact', 'receipt'],
  additionalProperties: false,
};
const pageIdentitySchema = {
  type: 'object',
  properties: {
    page: { type: 'integer', minimum: 1 },
    width: { type: 'number', exclusiveMinimum: 0 },
    height: { type: 'number', exclusiveMinimum: 0 },
    imageCount: { type: 'integer', minimum: 0 },
    wordCount: { type: 'integer', minimum: 0 },
    imageOnly: { type: 'boolean' },
  },
  required: ['page', 'width', 'height', 'imageCount', 'wordCount', 'imageOnly'],
  additionalProperties: false,
};
const inspectIdentitySchema = {
  type: 'object',
  properties: {
    format: { const: 'pdf' },
    pageCount: { type: 'integer', minimum: 1 },
    title: { type: ['string', 'null'] },
    author: { type: ['string', 'null'] },
    subject: { type: ['string', 'null'] },
    imageCount: { type: 'integer', minimum: 0 },
    wordCount: { type: 'integer', minimum: 0 },
    scannedPageCount: { type: 'integer', minimum: 0 },
    imageOnly: { type: 'boolean' },
    openingPages: { type: 'array', maxItems: 3, items: pageIdentitySchema },
  },
  required: [
    'format', 'pageCount', 'title', 'author', 'subject', 'imageCount', 'wordCount',
    'scannedPageCount', 'imageOnly', 'openingPages',
  ],
  additionalProperties: false,
};
const inspectOutputSchema = structuredClone(largeResultOutputSchema);
inspectOutputSchema.properties.identity = inspectIdentitySchema;
inspectOutputSchema.required.push('identity');

const definitions = new Map([
  ['pdf_inspect', {
    description: 'Inspect one current PDF revision. Always retain the complete observation at output and return a bounded identity containing page and document metadata without traversing document content.',
    outputSchema: inspectOutputSchema,
  }],
  ['pdf_extract_tables', {
    description: 'Extract tables from selected current PDF pages with deterministic published extraction. Set returnContent true to return complete bounded content; provide output to retain the complete immutable result. At least one result channel is required.',
    outputSchema: largeResultOutputSchema,
  }],
  ['pdf_find_table', {
    description: 'Find a caller-named table in one current PDF without deciding its business role. Set returnContent true to return complete bounded content; provide output to retain the complete immutable result. At least one result channel is required.',
    outputSchema: largeResultOutputSchema,
  }],
  ['pdf_ocr', {
    description: 'Observe selected current PDF pages through the fixed Aliyun qwen3.8-max OCR model and per-invocation Supen credential. Set returnContent true to return complete bounded content; provide output to retain the complete immutable result. At least one result channel is required.',
    outputSchema: largeResultOutputSchema,
  }],
  ['pdf_extract_table_details', {
    description: 'Read visual table cells, text spans, fonts, colors, and line evidence from selected current PDF pages. Set returnContent true to return complete bounded content; provide output to retain the complete immutable result. At least one result channel is required.',
    outputSchema: largeResultOutputSchema,
  }],
]);

const tools = manifest.tools.map((entry) => {
  const definition = definitions.get(entry.name);
  if (!definition) throw new Error(`pdf-provider-tool-definition-missing:${entry.name}`);
  const contractPath = path.join(pdfRoot, entry.inputContract.path);
  const bytes = readFileSync(contractPath);
  if (createHash('sha256').update(bytes).digest('hex') !== entry.inputContract.sha256) {
    throw new Error(`pdf-provider-input-contract-hash-invalid:${entry.name}`);
  }
  return {
    name: entry.name,
    description: definition.description,
    inputSchema: JSON.parse(bytes.toString('utf8')),
    outputSchema: definition.outputSchema,
    annotations: {
      readOnlyHint: true,
      idempotentHint: true,
      destructiveHint: false,
      openWorldHint: false,
    },
  };
});

async function callTool(name, args) {
  switch (name) {
    case 'pdf_inspect': return createToolResult(await pdfInspect(args));
    case 'pdf_extract_tables': return createToolResult(await pdfExtractTables(args));
    case 'pdf_find_table': return createToolResult(await pdfFindTable(args));
    case 'pdf_ocr': return createToolResult(await pdfOcr(args));
    case 'pdf_extract_table_details': return createToolResult(await pdfExtractTableDetails(args));
    default: throw Object.assign(new Error(`Unknown tool: ${name}`), { code: -32601 });
  }
}

async function pdfInspect(args) {
  rejectUnexpectedArgs(args, ['input', 'returnContent', 'output']);
  const input = path.resolve(requireString(args.input, 'input'));
  requireString(args.output, 'output');
  const result = await runJsonCandidateChain(pdfCandidates, ['inspect', input, '--json']);
  const payload = result.json;
  const document = payload?.document ?? payload;
  if (!document || !Number.isInteger(document.pages) || document.pages < 1) {
    throw new Error('pdf-inspect-runtime-result-invalid');
  }
  const metadata = document.metadata && typeof document.metadata === 'object' ? document.metadata : {};
  const identity = {
    format: 'pdf',
    pageCount: document.pages,
    title: nullableText(metadata.title),
    author: nullableText(metadata.author),
    subject: nullableText(metadata.subject),
    imageCount: nonnegativeInteger(document.image_count),
    wordCount: nonnegativeInteger(document.word_count),
    scannedPageCount: nonnegativeInteger(document.scanned_page_count),
    imageOnly: document.image_only === true,
    openingPages: (Array.isArray(document.page_sizes) ? document.page_sizes : []).slice(0, 3)
      .map((page) => ({
        page: positiveInteger(page.page),
        width: positiveNumber(page.width),
        height: positiveNumber(page.height),
        imageCount: nonnegativeInteger(page.image_count),
        wordCount: nonnegativeInteger(page.word_count),
        imageOnly: page.image_only === true,
      })),
  };
  return deliverLargeJsonResult({
    tool: 'pdf_inspect', args, runtime: commandRuntime(result), payload, sourcePaths: [input], identity,
  });
}

async function pdfExtractTables(args) {
  rejectUnexpectedArgs(args, ['input', 'pages', 'autoSpan', 'returnContent', 'output']);
  const input = path.resolve(requireString(args.input, 'input'));
  const commandArgs = ['extract-tables', input];
  appendPdfFlags(commandArgs, args);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return deliverLargeJsonResult({
    tool: 'pdf_extract_tables', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input],
  });
}

async function pdfFindTable(args) {
  rejectUnexpectedArgs(args, ['input', 'name', 'autoSpan', 'returnContent', 'output']);
  const input = path.resolve(requireString(args.input, 'input'));
  const commandArgs = ['find-table', input, requireString(args.name, 'name')];
  appendPdfFlags(commandArgs, args);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return deliverLargeJsonResult({
    tool: 'pdf_find_table', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input],
  });
}

async function pdfExtractTableDetails(args) {
  rejectUnexpectedArgs(args, ['input', 'pages', 'returnContent', 'output']);
  const input = path.resolve(requireString(args.input, 'input'));
  const commandArgs = ['extract-table-details', input];
  appendPages(commandArgs, args.pages);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return deliverLargeJsonResult({
    tool: 'pdf_extract_table_details', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input],
  });
}

async function pdfOcr(args) {
  rejectUnexpectedArgs(args, ['input', 'output', 'pages', 'returnContent']);
  const input = path.resolve(requireString(args.input, 'input'));
  const commandArgs = ['ocr', input, '--provider', 'llm', '--llm-model', 'qwen3.8-max'];
  appendPages(commandArgs, args.pages);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return deliverLargeJsonResult({
    tool: 'pdf_ocr', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input],
  });
}

function rejectUnexpectedArgs(args, allowed) {
  const unexpected = Object.keys(args).filter(key => !allowed.includes(key));
  if (unexpected.length > 0) {
    throw Object.assign(new Error(`Unexpected arguments: ${unexpected.join(', ')}`), { code: -32602 });
  }
}

function appendPdfFlags(commandArgs, args) {
  appendPages(commandArgs, args.pages);
  if (args.autoSpan === true) commandArgs.push('--auto-span');
}

function appendPages(commandArgs, pages) {
  if (Array.isArray(pages) && pages.length > 0) commandArgs.push('--pages', pages.join(','));
}

function commandRuntime(result) {
  return { command: result.command, arguments: result.args };
}

function nullableText(value) {
  return typeof value === 'string' && value.trim() ? value.trim() : null;
}

function nonnegativeInteger(value) {
  if (!Number.isInteger(value) || value < 0) throw new Error('pdf-inspect-runtime-result-invalid');
  return value;
}

function positiveInteger(value) {
  if (!Number.isInteger(value) || value < 1) throw new Error('pdf-inspect-runtime-result-invalid');
  return value;
}

function positiveNumber(value) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    throw new Error('pdf-inspect-runtime-result-invalid');
  }
  return value;
}

function readJson(file) {
  return JSON.parse(readFileSync(file, 'utf8'));
}

const server = new McpStdioServer({
  name: 'tiwater-pdf',
  version: packageJson.version,
  instructions: [
    'Generic PDF inspection, deterministic table extraction, and OCR fixed to Aliyun qwen3.8-max through the per-invocation Supen Gateway credential.',
    'Every output path is an immutable observation artifact identity: an identical request may replay identical bytes, while every different result uses a different path.',
    'PDF tools report technical observations only and never assign business roles, source dispositions, or delivery decisions.',
  ].join(' '),
  tools,
  callTool,
  logger: message => process.stderr.write(`${message}\n`),
});

server.start();
