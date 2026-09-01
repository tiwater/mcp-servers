#!/usr/bin/env node
import path from 'node:path';
import { McpStdioServer } from '../_shared/mcp-stdio.mjs';
import {
  commandCandidate,
  createToolResult,
  requireString,
  resolveRepoPath,
  runJsonCandidateChain,
} from '../_shared/tool-runtime.mjs';
import { deliverLargeJsonResult } from '../_shared/large-json-result.mjs';

const pdfPackageDir = resolveRepoPath('packages', 'pdf-cli');
const pdfModulePath = resolveRepoPath('packages', 'pdf-cli');

const pdfCandidates = [
  commandCandidate('tiwater-pdf'),
  commandCandidate('python3', ['-m', 'tiwater_pdf.cli'], {
    cwd: pdfPackageDir,
    env: {
      PYTHONPATH: [pdfModulePath, process.env.PYTHONPATH || ''].filter(Boolean).join(path.delimiter),
    },
  }),
];

const artifactSchema = {
  type: 'object',
  properties: {
    path: { type: 'string' },
    sha256: { type: 'string', pattern: '^[0-9a-f]{64}$' },
    bytes: { type: 'integer', minimum: 0 },
  },
  required: ['path', 'sha256', 'bytes'],
  additionalProperties: false,
};
const largeResultOutputSchema = {
  type: 'object',
  properties: {
    tool: { type: 'string' },
    runtime: { type: 'string' },
    sources: { type: 'array', minItems: 1, maxItems: 1, items: artifactSchema },
    returnContent: { type: 'boolean' },
    artifact: { anyOf: [artifactSchema, { type: 'null' }] },
    receipt: {
      type: 'object',
      properties: {
        contentBytes: { type: 'integer', minimum: 0 },
        contentReturned: { type: 'boolean' },
        contentWritten: { type: 'boolean' },
      },
      required: ['contentBytes', 'contentReturned', 'contentWritten'],
      additionalProperties: false,
    },
    content: {},
  },
  required: ['tool', 'runtime', 'sources', 'returnContent', 'artifact', 'receipt'],
  additionalProperties: false,
};

function resultProperties(extra = {}) {
  return {
    input: { type: 'string', minLength: 1, 'x-tiwater-file-role': 'read' },
    returnContent: {
      type: 'boolean',
      description: 'Return the complete result directly when it fits the published response limit.',
    },
    output: {
      type: 'string',
      minLength: 1,
      description: 'Absolute path for a new JSON file containing the complete result. Existing files are never overwritten.',
      'x-tiwater-file-role': 'write',
    },
    ...extra,
  };
}

const tools = [
  {
    name: 'pdf_inspect',
    description: 'Inspect a PDF and produce metadata and page count. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: {
      type: 'object',
      properties: resultProperties(),
      required: ['input', 'returnContent'],
      additionalProperties: false,
    },
    outputSchema: largeResultOutputSchema,
  },
  {
    name: 'pdf_extract_tables',
    description: 'Extract tables from a PDF with deterministic published extraction. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: {
      type: 'object',
      properties: resultProperties({
        pages: { type: 'array', items: { type: 'number' } },
        autoSpan: { type: 'boolean' },
      }),
      required: ['input', 'returnContent'],
      additionalProperties: false,
    },
    outputSchema: largeResultOutputSchema,
  },
  {
    name: 'pdf_find_table',
    description: 'Find a named table in a PDF and produce the matched table data. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: {
      type: 'object',
      properties: resultProperties({
        name: { type: 'string' },
        autoSpan: { type: 'boolean' },
      }),
      required: ['input', 'name', 'returnContent'],
      additionalProperties: false,
    },
    outputSchema: largeResultOutputSchema,
  },
  {
    name: 'pdf_ocr',
    description: 'OCR current PDF pages through the per-invocation Supen Gateway using only the pinned Aliyun qwen3.8-max vision model. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: {
      type: 'object',
      properties: resultProperties({
        pages: { type: 'array', items: { type: 'number' } },
      }),
      required: ['input', 'returnContent'],
      additionalProperties: false,
    },
    outputSchema: largeResultOutputSchema,
  },
  {
    name: 'pdf_extract_table_details',
    description: 'Extract detected PDF tables with visual cell boxes, text spans, colors, fonts, and line evidence for format validation. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: {
      type: 'object',
      properties: resultProperties({
        pages: { type: 'array', items: { type: 'number' } },
      }),
      required: ['input', 'returnContent'],
      additionalProperties: false,
    },
    outputSchema: largeResultOutputSchema,
  },
];

async function callTool(name, args) {
  switch (name) {
    case 'pdf_inspect':
      return createToolResult(await pdfInspect(args));
    case 'pdf_extract_tables':
      return createToolResult(await pdfExtractTables(args));
    case 'pdf_find_table':
      return createToolResult(await pdfFindTable(args));
    case 'pdf_ocr':
      return createToolResult(await pdfOcr(args));
    case 'pdf_extract_table_details':
      return createToolResult(await pdfExtractTableDetails(args));
    default:
      throw Object.assign(new Error(`Unknown tool: ${name}`), { code: -32601 });
  }
}

async function pdfInspect(args) {
  rejectUnexpectedArgs(args, ['input', 'returnContent', 'output']);
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(pdfCandidates, ['inspect', input, '--json']);
  return deliverLargeJsonResult({ tool: 'pdf_inspect', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

async function pdfExtractTables(args) {
  rejectUnexpectedArgs(args, ['input', 'pages', 'autoSpan', 'returnContent', 'output']);
  const input = path.resolve(requireString(args.input, 'input'));
  const commandArgs = ['extract-tables', input];
  appendPdfFlags(commandArgs, args);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return deliverLargeJsonResult({ tool: 'pdf_extract_tables', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

async function pdfFindTable(args) {
  rejectUnexpectedArgs(args, ['input', 'name', 'autoSpan', 'returnContent', 'output']);
  const input = path.resolve(requireString(args.input, 'input'));
  const name = requireString(args.name, 'name');
  const commandArgs = ['find-table', input, name];
  appendPdfFlags(commandArgs, args);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return deliverLargeJsonResult({ tool: 'pdf_find_table', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

async function pdfExtractTableDetails(args) {
  rejectUnexpectedArgs(args, ['input', 'pages', 'returnContent', 'output']);
  const input = path.resolve(requireString(args.input, 'input'));
  const commandArgs = ['extract-table-details', input];
  if (Array.isArray(args.pages) && args.pages.length > 0) {
    commandArgs.push('--pages', args.pages.join(','));
  }
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return deliverLargeJsonResult({ tool: 'pdf_extract_table_details', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

async function pdfOcr(args) {
  rejectUnexpectedArgs(args, ['input', 'output', 'pages', 'returnContent']);
  const input = path.resolve(requireString(args.input, 'input'));
  const commandArgs = ['ocr', input, '--provider', 'llm', '--llm-model', 'qwen3.8-max'];
  if (Array.isArray(args.pages) && args.pages.length > 0) commandArgs.push('--pages', args.pages.join(','));
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return deliverLargeJsonResult({ tool: 'pdf_ocr', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

function rejectUnexpectedArgs(args, allowed) {
  const unexpected = Object.keys(args).filter(key => !allowed.includes(key));
  if (unexpected.length > 0) {
    throw Object.assign(new Error(`Unexpected arguments: ${unexpected.join(', ')}`), { code: -32602 });
  }
}

function appendPdfFlags(commandArgs, args) {
  if (Array.isArray(args.pages) && args.pages.length > 0) {
    commandArgs.push('--pages', args.pages.join(','));
  }
  if (args.autoSpan) commandArgs.push('--auto-span');
}

function commandRuntime(result) {
  return `${result.command} ${result.args.join(' ')}`;
}

const server = new McpStdioServer({
  name: 'tiwater-pdf',
  version: '0.23.0',
  instructions: 'Generic PDF inspection, deterministic table extraction, and OCR fixed to Aliyun qwen3.8-max through the per-invocation Supen Gateway credential.',
  tools,
  callTool,
  logger: message => process.stderr.write(`${message}\n`),
});

server.start();
