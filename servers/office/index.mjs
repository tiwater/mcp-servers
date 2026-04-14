#!/usr/bin/env node
import path from 'node:path';
import { McpStdioServer } from '../_shared/mcp-stdio.mjs';
import {
  commandCandidate,
  createToolResult,
  maybeReadJson,
  requireString,
  resolveRepoPath,
  runCandidateChain,
  runJsonCandidateChain,
  withTempJsonFile,
} from '../_shared/tool-runtime.mjs';

const docxProject = resolveRepoPath('packages', 'docx-cli', 'docx.csproj');
const xlsxProject = resolveRepoPath('packages', 'xlsx-cli', 'xlsx.csproj');

const docxCandidates = [
  commandCandidate('tiwater-docx'),
  commandCandidate('dotnet', ['run', '--project', docxProject, '--']),
];

const xlsxCandidates = [
  commandCandidate('tiwater-xlsx'),
  commandCandidate('dotnet', ['run', '--project', xlsxProject, '--']),
];

const tools = [
  {
    name: 'docx_inspect',
    description: 'Inspect a DOCX document and return structural and formatting metrics.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string', description: 'Absolute or relative path to a .docx file.' } },
      required: ['input'],
    },
  },
  {
    name: 'docx_compare',
    description: 'Compare two DOCX files and report package, metric, and style differences.',
    inputSchema: {
      type: 'object',
      properties: {
        baseline: { type: 'string' },
        updated: { type: 'string' },
      },
      required: ['baseline', 'updated'],
    },
  },
  {
    name: 'docx_validate_template_transform',
    description: 'Validate whether a source DOCX template and target DOCX template are structurally compatible.',
    inputSchema: {
      type: 'object',
      properties: {
        sourceTemplate: { type: 'string' },
        targetTemplate: { type: 'string' },
      },
      required: ['sourceTemplate', 'targetTemplate'],
    },
  },
  {
    name: 'docx_strip_direct_formatting',
    description: 'Copy a DOCX and remove direct paragraph and run formatting while preserving styles.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
      },
      required: ['input', 'output'],
    },
  },
  {
    name: 'docx_replace_style_ids',
    description: 'Copy a DOCX and replace style IDs based on a provided style map object or JSON file.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
        styleMap: { type: 'object', additionalProperties: { type: 'string' } },
        styleMapPath: { type: 'string' },
      },
      required: ['input', 'output'],
    },
  },
  {
    name: 'docx_export_json',
    description: 'Export the body content of a DOCX document as structured JSON.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
      },
      required: ['input'],
    },
  },
  {
    name: 'docx_fill_template',
    description: 'Fill DOCX placeholders using a data object or an existing JSON data file.',
    inputSchema: {
      type: 'object',
      properties: {
        template: { type: 'string' },
        output: { type: 'string' },
        data: { type: 'object' },
        dataPath: { type: 'string' },
      },
      required: ['template', 'output'],
    },
  },
  {
    name: 'xlsx_inspect',
    description: 'Inspect an XLSX workbook and return sheet-level metrics.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string' } },
      required: ['input'],
    },
  },
  {
    name: 'xlsx_export_json',
    description: 'Export workbook sheet data from XLSX as structured JSON.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
      },
      required: ['input'],
    },
  },
  {
    name: 'xlsx_fill_template',
    description: 'Fill an XLSX template using a data object or an existing JSON data file.',
    inputSchema: {
      type: 'object',
      properties: {
        template: { type: 'string' },
        output: { type: 'string' },
        data: { type: 'object' },
        dataPath: { type: 'string' },
      },
      required: ['template', 'output'],
    },
  },
];

async function callTool(name, args) {
  switch (name) {
    case 'docx_inspect':
      return createToolResult(await docxInspect(args));
    case 'docx_compare':
      return createToolResult(await docxCompare(args));
    case 'docx_validate_template_transform':
      return createToolResult(await docxValidateTemplateTransform(args));
    case 'docx_strip_direct_formatting':
      return createToolResult(await docxStripDirectFormatting(args));
    case 'docx_replace_style_ids':
      return createToolResult(await docxReplaceStyleIds(args));
    case 'docx_export_json':
      return createToolResult(await docxExportJson(args));
    case 'docx_fill_template':
      return createToolResult(await docxFillTemplate(args));
    case 'xlsx_inspect':
      return createToolResult(await xlsxInspect(args));
    case 'xlsx_export_json':
      return createToolResult(await xlsxExportJson(args));
    case 'xlsx_fill_template':
      return createToolResult(await xlsxFillTemplate(args));
    default:
      throw Object.assign(new Error(`Unknown tool: ${name}`), { code: -32601 });
  }
}

async function docxInspect(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(docxCandidates, ['inspect', input, '--json']);
  return { tool: 'docx_inspect', runtime: commandRuntime(result), report: result.json };
}

async function docxCompare(args) {
  const baseline = requireString(args.baseline, 'baseline');
  const updated = requireString(args.updated, 'updated');
  const result = await runJsonCandidateChain(docxCandidates, ['compare', baseline, updated, '--json']);
  return { tool: 'docx_compare', runtime: commandRuntime(result), report: result.json };
}

async function docxValidateTemplateTransform(args) {
  const sourceTemplate = requireString(args.sourceTemplate, 'sourceTemplate');
  const targetTemplate = requireString(args.targetTemplate, 'targetTemplate');
  const result = await runJsonCandidateChain(docxCandidates, ['validate-template-transform', sourceTemplate, targetTemplate, '--json']);
  return { tool: 'docx_validate_template_transform', runtime: commandRuntime(result), report: result.json };
}

async function docxStripDirectFormatting(args) {
  const input = requireString(args.input, 'input');
  const output = requireString(args.output, 'output');
  const result = await runCandidateChain(docxCandidates, ['strip-direct-formatting', input, output]);
  return { tool: 'docx_strip_direct_formatting', runtime: commandRuntime(result), outputPath: output, stdout: result.stdout.trim() };
}

async function docxReplaceStyleIds(args) {
  const input = requireString(args.input, 'input');
  const output = requireString(args.output, 'output');
  if (args.styleMapPath) {
    const styleMapPath = requireString(args.styleMapPath, 'styleMapPath');
    const result = await runCandidateChain(docxCandidates, ['replace-style-ids', input, output, styleMapPath]);
    return { tool: 'docx_replace_style_ids', runtime: commandRuntime(result), outputPath: output, stdout: result.stdout.trim(), styleMapPath };
  }
  if (!args.styleMap || typeof args.styleMap !== 'object' || Array.isArray(args.styleMap)) {
    throw Object.assign(new Error('styleMap or styleMapPath is required'), { code: -32602 });
  }
  return withTempJsonFile(args.styleMap, async styleMapPath => {
    const result = await runCandidateChain(docxCandidates, ['replace-style-ids', input, output, styleMapPath]);
    return { tool: 'docx_replace_style_ids', runtime: commandRuntime(result), outputPath: output, stdout: result.stdout.trim() };
  });
}

async function docxExportJson(args) {
  const input = requireString(args.input, 'input');
  if (args.output) {
    const output = requireString(args.output, 'output');
    const result = await runCandidateChain(docxCandidates, ['export-json', input, output]);
    return { tool: 'docx_export_json', runtime: commandRuntime(result), outputPath: output, document: await maybeReadJson(output) };
  }
  const result = await runCandidateChain(docxCandidates, ['export-json', input]);
  return { tool: 'docx_export_json', runtime: commandRuntime(result), document: JSON.parse(result.stdout) };
}

async function docxFillTemplate(args) {
  const template = requireString(args.template, 'template');
  const output = requireString(args.output, 'output');
  if (args.dataPath) {
    const dataPath = requireString(args.dataPath, 'dataPath');
    const result = await runCandidateChain(docxCandidates, ['fill-template', template, dataPath, output]);
    return { tool: 'docx_fill_template', runtime: commandRuntime(result), outputPath: output, stdout: result.stdout.trim() };
  }
  if (args.data === undefined) {
    throw Object.assign(new Error('data or dataPath is required'), { code: -32602 });
  }
  return withTempJsonFile(args.data, async dataPath => {
    const result = await runCandidateChain(docxCandidates, ['fill-template', template, dataPath, output]);
    return { tool: 'docx_fill_template', runtime: commandRuntime(result), outputPath: output, stdout: result.stdout.trim() };
  });
}

async function xlsxInspect(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(xlsxCandidates, ['inspect', input, '--json']);
  return { tool: 'xlsx_inspect', runtime: commandRuntime(result), report: result.json };
}

async function xlsxExportJson(args) {
  const input = requireString(args.input, 'input');
  if (args.output) {
    const output = requireString(args.output, 'output');
    const result = await runCandidateChain(xlsxCandidates, ['export-json', input, output]);
    return { tool: 'xlsx_export_json', runtime: commandRuntime(result), outputPath: output, workbook: await maybeReadJson(output) };
  }
  const result = await runCandidateChain(xlsxCandidates, ['export-json', input]);
  return { tool: 'xlsx_export_json', runtime: commandRuntime(result), workbook: JSON.parse(result.stdout) };
}

async function xlsxFillTemplate(args) {
  const template = requireString(args.template, 'template');
  const output = requireString(args.output, 'output');
  if (args.dataPath) {
    const dataPath = requireString(args.dataPath, 'dataPath');
    const result = await runCandidateChain(xlsxCandidates, ['fill-template', template, dataPath, output]);
    return { tool: 'xlsx_fill_template', runtime: commandRuntime(result), outputPath: output, stdout: result.stdout.trim() };
  }
  if (args.data === undefined) {
    throw Object.assign(new Error('data or dataPath is required'), { code: -32602 });
  }
  return withTempJsonFile(args.data, async dataPath => {
    const result = await runCandidateChain(xlsxCandidates, ['fill-template', template, dataPath, output]);
    return { tool: 'xlsx_fill_template', runtime: commandRuntime(result), outputPath: output, stdout: result.stdout.trim() };
  });
}

function commandRuntime(result) {
  return `${result.command} ${result.args.join(' ')}`;
}

const server = new McpStdioServer({
  name: 'tiwater-office',
  version: '0.1.2',
  instructions: 'Shared Office MCP server for DOCX and XLSX inspection, export, validation, and template filling.',
  tools,
  callTool,
  logger: message => process.stderr.write(`${message}\n`),
});

server.start();
