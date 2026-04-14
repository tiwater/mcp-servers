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
    description: 'Extract tables from a PDF, optionally limiting pages or enabling multi-page spanning and LLM fallback.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        pages: { type: 'array', items: { type: 'number' } },
        autoSpan: { type: 'boolean' },
        llmFallback: { type: 'boolean' },
        apiKey: { type: 'string' },
        llmModel: { type: 'string' },
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
        llmFallback: { type: 'boolean' },
        apiKey: { type: 'string' },
        llmModel: { type: 'string' },
      },
      required: ['input', 'name'],
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
  const input = requireString(args.input, 'input');
  const commandArgs = ['extract-tables', input];
  appendPdfFlags(commandArgs, args);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return { tool: 'pdf_extract_tables', runtime: commandRuntime(result), report: result.json };
}

async function pdfFindTable(args) {
  const input = requireString(args.input, 'input');
  const name = requireString(args.name, 'name');
  const commandArgs = ['find-table', input, name];
  appendPdfFlags(commandArgs, args);
  commandArgs.push('--json');
  const result = await runJsonCandidateChain(pdfCandidates, commandArgs);
  return { tool: 'pdf_find_table', runtime: commandRuntime(result), report: result.json };
}

function appendPdfFlags(commandArgs, args) {
  if (Array.isArray(args.pages) && args.pages.length > 0) {
    commandArgs.push('--pages', args.pages.join(','));
  }
  if (args.autoSpan) commandArgs.push('--auto-span');
  if (args.llmFallback) commandArgs.push('--llm-fallback');
  if (typeof args.apiKey === 'string' && args.apiKey) commandArgs.push('--api-key', args.apiKey);
  if (typeof args.llmModel === 'string' && args.llmModel) commandArgs.push('--llm-model', args.llmModel);
}

function commandRuntime(result) {
  return `${result.command} ${result.args.join(' ')}`;
}

const server = new McpStdioServer({
  name: 'tiwater-pdf',
  version: '0.1.2',
  instructions: 'Shared PDF MCP server for inspection and table extraction.',
  tools,
  callTool,
  logger: message => process.stderr.write(`${message}\n`),
});

server.start();
