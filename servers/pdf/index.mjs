#!/usr/bin/env node
import path from 'node:path';
import { createHash } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { McpStdioServer } from '../_shared/mcp-stdio.mjs';
import {
  commandCandidate,
  createToolResult,
  requireString,
  resolveRepoPath,
  runJsonCandidateChain,
} from '../_shared/tool-runtime.mjs';

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

const tools = [
  {
    name: 'pdf_inspect',
    description: 'Inspect a PDF and return metadata and page count.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string' } },
      required: ['input'],
    },
  },
  {
    name: 'pdf_extract_tables',
    description: 'Extract tables from a PDF with deterministic published extraction.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        pages: { type: 'array', items: { type: 'number' } },
        autoSpan: { type: 'boolean' },
      },
      required: ['input'],
    },
  },
  {
    name: 'pdf_find_table',
    description: 'Find a named table in a PDF and return the matched table data.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        name: { type: 'string' },
        autoSpan: { type: 'boolean' },
      },
      required: ['input', 'name'],
    },
  },
  {
    name: 'pdf_ocr',
    description: 'OCR current PDF pages through the per-invocation Supen Gateway using the pinned Aliyun qwen3.8-max vision model and write a JSON evidence artifact.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
        pages: { type: 'array', items: { type: 'number' } },
      },
      required: ['input', 'output'],
      additionalProperties: false,
    },
  },
  {
    name: 'pdf_extract_table_details',
    description: 'Extract detected PDF tables with visual cell bboxes, text spans, colors, fonts, and line evidence for format validation.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        pages: { type: 'array', items: { type: 'number' } },
      },
      required: ['input'],
    },
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
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(pdfCandidates, ['inspect', input, '--json']);
  return { tool: 'pdf_inspect', runtime: commandRuntime(result), report: result.json };
}

async function pdfExtractTables(args) {
  rejectUnexpectedArgs(args, ['input', 'pages', 'autoSpan']);
  const input = requireString(args.input, 'input');
  const commandArgs = ['extract-tables', input];
  appendPdfFlags(commandArgs, args);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return { tool: 'pdf_extract_tables', runtime: commandRuntime(result), report: result.json };
}

async function pdfFindTable(args) {
  rejectUnexpectedArgs(args, ['input', 'name', 'autoSpan']);
  const input = requireString(args.input, 'input');
  const name = requireString(args.name, 'name');
  const commandArgs = ['find-table', input, name];
  appendPdfFlags(commandArgs, args);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return { tool: 'pdf_find_table', runtime: commandRuntime(result), report: result.json };
}

async function pdfExtractTableDetails(args) {
  const input = requireString(args.input, 'input');
  const commandArgs = ['extract-table-details', input];
  if (Array.isArray(args.pages) && args.pages.length > 0) {
    commandArgs.push('--pages', args.pages.join(','));
  }
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return { tool: 'pdf_extract_table_details', runtime: commandRuntime(result), report: result.json };
}

async function pdfOcr(args) {
  rejectUnexpectedArgs(args, ['input', 'output', 'pages']);
  const input = requireString(args.input, 'input');
  const output = path.resolve(requireString(args.output, 'output'));
  const commandArgs = ['ocr', input, '--provider', 'llm', '--llm-model', 'qwen3.8-max'];
  if (Array.isArray(args.pages) && args.pages.length > 0) commandArgs.push('--pages', args.pages.join(','));
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  const bytes = Buffer.from(`${JSON.stringify(result.json, null, 2)}\n`, 'utf8');
  await mkdir(path.dirname(output), { recursive: true });
  await writeFile(output, bytes, { flag: 'wx' });
  return {
    tool: 'pdf_ocr',
    runtime: commandRuntime(result),
    artifact: { path: output, sha256: createHash('sha256').update(bytes).digest('hex'), bytes: bytes.length },
    summary: { model: 'qwen3.8-max', pageCount: result.json?.page_count ?? result.json?.pages?.length ?? 0 },
  };
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
  version: '0.22.2',
  instructions: 'Generic PDF inspection, deterministic table extraction, and OCR fixed to Aliyun qwen3.8-max through the per-invocation Supen Gateway credential.',
  tools,
  callTool,
  logger: message => process.stderr.write(`${message}\n`),
});

server.start();
