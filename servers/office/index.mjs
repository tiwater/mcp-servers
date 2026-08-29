#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { createReadStream } from 'node:fs';
import { mkdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import { McpServer } from '@modelcontextprotocol/server';
import { serveStdio } from '@modelcontextprotocol/server/stdio';
import * as z from 'zod/v4';
import {
  commandCandidate,
  createToolResult,
  requireString,
  runCandidateChain,
  runJsonCandidateChain,
  withTempJsonFile,
} from '../_shared/tool-runtime.mjs';

const packageMetadata = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
const inputContractManifest = JSON.parse(await readFile(
  new URL('./contracts/tiwater-office-provider-contract-manifest-v1.json', import.meta.url),
  'utf8',
));
if (inputContractManifest.provider?.id !== packageMetadata.name
    || inputContractManifest.provider?.version !== packageMetadata.version) {
  throw new Error('Office MCP input contract manifest does not match the installed provider package');
}
const inputContracts = new Map(await Promise.all(inputContractManifest.tools.map(async entry => {
  const schema = JSON.parse(await readFile(new URL(`./contracts/${entry.name}.schema.json`, import.meta.url), 'utf8'));
  return [entry.name, z.fromJSONSchema(schema)];
})));
const invocationCwd = process.cwd();

function inputContract(toolName) {
  const contract = inputContracts.get(toolName);
  if (!contract) throw new Error(`Missing provider-owned MCP input contract: ${toolName}`);
  return contract;
}

const docxCandidates = [
  commandCandidate('tiwater-docx', [], { cwd: invocationCwd }),
];

const xlsxCandidates = [
  commandCandidate('tiwater-xlsx', [], { cwd: invocationCwd }),
];

const pptxCandidates = [
  commandCandidate('tiwater-pptx', [], { cwd: invocationCwd }),
];

const convertCandidates = [
  commandCandidate('tiwater-convert', [], { cwd: invocationCwd }),
];

const runtimeIdentity = z.object({
  command: z.string(),
  cwd: z.string(),
}).strict();

const artifact = z.object({
  path: z.string(),
  sha256: z.string().regex(/^[0-9a-f]{64}$/),
  bytes: z.number().int().nonnegative(),
}).strict();

const renderFileIdentity = z.object({
  sha256: z.string().regex(/^[a-f0-9]{64}$/),
  size_bytes: z.number().int().positive(),
}).strict();
const nativeRenderReceipt = z.object({
  status: z.literal('ok'),
  input: z.string(),
  input_sha256: z.string().regex(/^[a-f0-9]{64}$/),
  output: z.string(),
  output_sha256: z.string().regex(/^[a-f0-9]{64}$/),
  source_format: z.enum(['doc', 'docx', 'odt', 'rtf', 'xls', 'xlsx', 'ods', 'ppt', 'pptx', 'odp']),
  target_format: z.literal('pdf'),
  version: z.string(),
  backend: z.enum(['wps', 'et', 'wpp']),
  fallback_reason: z.null(),
  page_count: z.number().int().positive(),
  native_render_provenance: z.object({
    schema: z.literal('tiwater.convert-native-render-provenance/v1'),
    backend: z.enum(['wps', 'et', 'wpp']),
    input: renderFileIdentity,
    output: renderFileIdentity,
    page_count: z.number().int().positive(),
  }).passthrough(),
}).passthrough();
const nativeRenderOutput = z.object({
  tool: z.literal('office_render_pdf'),
  runtime: runtimeIdentity,
  pdf: artifact,
  receipt: artifact,
  summary: z.object({
    sourceFormat: z.string(),
    backend: z.enum(['wps', 'et', 'wpp']),
    pageCount: z.number().int().positive(),
  }).strict(),
}).strict();

const xlsxFixedTools = [
  {"name":"xlsx_set_cell_value","description":"Set current workbook cell values."},
  {"name":"xlsx_set_cell_number_format","description":"Set current workbook cell number formats."},
  {"name":"xlsx_set_rich_text_cell_value","description":"Set current workbook rich-text cell values."},
  {"name":"xlsx_set_range_values","description":"Set rectangular values in a current workbook."},
  {"name":"xlsx_insert_rows","description":"Insert rows into a current worksheet."},
  {"name":"xlsx_delete_rows","description":"Structurally delete rows from a current worksheet."},
  {"name":"xlsx_copy_row","description":"Copy current worksheet rows."},
  {"name":"xlsx_expand_section_rows","description":"Expand current worksheet row sections from visible anchors."},
  {"name":"xlsx_set_print_area","description":"Set current worksheet print areas."},
  {"name":"xlsx_set_page_setup","description":"Set current worksheet page properties."},
  {"name":"xlsx_set_row_page_breaks","description":"Set current worksheet row page breaks."},
  {"name":"xlsx_set_column_width","description":"Set current worksheet column widths."},
];

function fixedToolDefinitions(definitions) {
  return definitions.map(definition => ({
    name: definition.name,
    description: definition.description,
    inputSchema: inputContract(definition.name),
    outputSchema: fixedEditOutput(definition.name),
    handler: args => fixedEdit(definition.name, args),
  }));
}

function fixedEditOutput(tool) {
  return z.object({
    tool: z.literal(tool), runtime: runtimeIdentity, receipt: artifact, output: artifact.nullable(),
    summary: z.object({ pass: z.boolean(), operationCount: z.number().int().nonnegative(), appliedCount: z.number().int().nonnegative() }).strict(),
  }).strict();
}

function artifactOutput(tool) {
  return z.object({ tool: z.literal(tool), runtime: runtimeIdentity, source: artifact, artifact }).strict();
}

const tools = [
  {
    name: 'docx_inspect',
    description: 'Write the default single full-document DOCX observation containing body content, structure, placeholders, comments, anchors, tables, fields, flow, fonts, and formatting metrics. Its tables section includes the current revision and native refs for every observed table, row, cell, paragraph, and run; use those refs directly for mutation instead of listing the same descendants again. Do not also call docx_export_json for the same observation.',
    inputSchema: inputContract('docx_inspect'),
    outputSchema: artifactOutput('docx_inspect'),
    handler: docxInspect,
  },
  {
    name: 'docx_inspect_tables',
    description: 'Inspect current DOCX tables, cells, merges, paragraphs, runs, and formatting, returning the current revision and native refs for every observed table, row, cell, paragraph, and run. Use those refs directly for mutation instead of listing the same descendants again.',
    inputSchema: inputContract('docx_inspect_tables'),
    outputSchema: artifactOutput('docx_inspect_tables'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxInspectTables,
  },
  {
    name: 'docx_list_objects',
    description: 'List one small page of revision-bound DOCX object identities for structure discovery, not selection by text. To select a table or row by any descendant cell text, call docx_find_literal with kind table or row instead of listing the document. After selecting a table or row, call docx_read_object once to obtain every descendant ref; do not list its rows and cells separately. Use parentRef only when nearest-child paging itself is the requested observation.',
    inputSchema: inputContract('docx_list_objects'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: args => docxObservation('docx_list_objects', args),
  },
  {
    name: 'docx_find_literal',
    description: 'Find exact current text in one small page of revision-bound native DOCX objects. With kind table or row, matching includes all descendant cell text. After selecting a table or row, call docx_read_object once to obtain every descendant ref; do not list its rows and cells separately.',
    inputSchema: inputContract('docx_find_literal'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: args => docxObservation('docx_find_literal', args),
  },
  {
    name: 'docx_read_object',
    description: 'Write one selected revision-bound native DOCX object, its Open XML, and every descendant object ref to a new JSON artifact. One table read provides all row, cell, paragraph, run, text, and drawing refs needed for mutation; do not enumerate those descendants with repeated list calls.',
    inputSchema: inputContract('docx_read_object'),
    outputSchema: artifactOutput('docx_read_object'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxReadObject,
  },
  {
    name: 'docx_copy_content',
    description: 'Copy selected current DOCX content directly into selected current target cells while retaining target cell formatting and source inline meaning.',
    inputSchema: inputContract('docx_copy_content'),
    outputSchema: fixedEditOutput('docx_copy_content'),
    handler: args => fixedEdit('docx_copy_content', args),
  },
  {
    name: 'docx_copy_object',
    description: 'Copy selected current DOCX objects directly under a selected current target parent object.',
    inputSchema: inputContract('docx_copy_object'),
    outputSchema: fixedEditOutput('docx_copy_object'),
    handler: args => fixedEdit('docx_copy_object', args),
  },
  {
    name: 'docx_delete_object',
    description: 'Delete selected current DOCX objects directly from the current target document.',
    inputSchema: inputContract('docx_delete_object'),
    outputSchema: fixedEditOutput('docx_delete_object'),
    handler: args => fixedEdit('docx_delete_object', args),
  },
  {
    name: 'docx_merge_cells',
    description: 'Merge selected current DOCX cells when they form one closed rectangle.',
    inputSchema: inputContract('docx_merge_cells'),
    outputSchema: fixedEditOutput('docx_merge_cells'),
    handler: args => fixedEdit('docx_merge_cells', args),
  },
  {
    name: 'docx_split_cells',
    description: 'Split selected current DOCX merged cells.',
    inputSchema: inputContract('docx_split_cells'),
    outputSchema: fixedEditOutput('docx_split_cells'),
    handler: args => fixedEdit('docx_split_cells', args),
  },
  {
    name: 'docx_compare',
    description: 'Compare two DOCX files and report package, metric, and style differences.',
    inputSchema: inputContract('docx_compare'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxCompare,
  },
  {
    name: 'docx_export_json',
    description: 'Write a body-only DOCX JSON projection only when a downstream consumer explicitly requires that format. It is not a companion to docx_inspect and does not replace bounded list/find/read object selection.',
    inputSchema: inputContract('docx_export_json'),
    outputSchema: artifactOutput('docx_export_json'),
    handler: docxExportJson,
  },
  {
    name: 'docx_validate',
    description: 'Validate a current DOCX package against the published OpenXML contract.',
    inputSchema: inputContract('docx_validate'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxValidate,
  },
  {
    name: 'docx_validate_font_policy',
    description: 'Validate current DOCX text against an explicit font policy.',
    inputSchema: inputContract('docx_validate_font_policy'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxValidateFontPolicy,
  },
  {
    name: 'docx_validate_toc_style_policy',
    description: 'Validate current DOCX table-of-contents paragraph styles against an explicit policy.',
    inputSchema: inputContract('docx_validate_toc_style_policy'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxValidateTocStylePolicy,
  },
  {
    name: 'docx_strip_direct_formatting',
    description: 'Remove direct paragraph and run formatting while preserving styles.',
    inputSchema: inputContract('docx_strip_direct_formatting'),
    handler: docxStripDirectFormatting,
  },
  {
    name: 'docx_replace_style_ids',
    description: 'Replace current DOCX style IDs from an explicit style map.',
    inputSchema: inputContract('docx_replace_style_ids'),
    handler: docxReplaceStyleIds,
  },
  {
    name: 'office_render_pdf',
    description: 'Render a current Office document to PDF with its required native WPS backend and write the complete provider receipt as evidence. The input extension selects Writer, Spreadsheets, or Presentation; fallback rendering is rejected.',
    inputSchema: inputContract('office_render_pdf'),
    outputSchema: nativeRenderOutput,
    handler: officeRenderPdf,
  },
  {
    name: 'xlsx_convert_legacy',
    description: 'Convert a current legacy XLS workbook to XLSX using the published native ET backend.',
    inputSchema: inputContract('xlsx_convert_legacy'),
    handler: xlsxConvertLegacy,
  },
  {
    name: 'xlsx_inspect',
    description: 'Inspect a current XLSX workbook or legacy XLS workbook and write one JSON observation containing workbook structure, exported values, formulas, styles, merged ranges, and any published legacy-format conversion evidence.',
    inputSchema: inputContract('xlsx_inspect'),
    outputSchema: artifactOutput('xlsx_inspect'),
    handler: xlsxInspect,
  },
  {
    name: 'xlsx_export_json',
    description: 'Export workbook sheet data from XLSX as structured JSON.',
    inputSchema: inputContract('xlsx_export_json'),
    outputSchema: artifactOutput('xlsx_export_json'),
    handler: xlsxExportJson,
  },
  ...fixedToolDefinitions(xlsxFixedTools),
  {
    name: 'xlsx_validate',
    description: 'Validate an XLSX workbook package and return Open XML validation evidence.',
    inputSchema: inputContract('xlsx_validate'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: xlsxValidate,
  },
  {
    name: 'pptx_inspect',
    description: 'Inspect a PPTX file and write one detailed JSON observation containing slides, masters, layouts, shapes, transforms, paragraphs, runs, and placeholders.',
    inputSchema: inputContract('pptx_inspect'),
    outputSchema: artifactOutput('pptx_inspect'),
    handler: pptxInspect,
  },
  {
    name: 'pptx_export_json',
    description: 'Export PPTX slide text, notes, and placeholder hints to a new JSON artifact without returning the full presentation through MCP.',
    inputSchema: inputContract('pptx_export_json'),
    outputSchema: artifactOutput('pptx_export_json'),
    handler: pptxExportJson,
  },
  {
    name: 'pptx_apply_template',
    description: 'Apply one deterministic PPTX template-application plan to a current presentation. This tool executes the published plan; it does not select a template or derive business content, slide mappings, geometry, or formatting decisions.',
    inputSchema: inputContract('pptx_apply_template'),
    outputSchema: fixedEditOutput('pptx_apply_template'),
    handler: args => fixedEdit('pptx_apply_template', args),
  },
  {
    name: 'pptx_apply_format',
    description: 'Apply one deterministic PPTX formatting plan to a current presentation. This tool executes published formatting operations; it does not derive values, coordinates, or business decisions.',
    inputSchema: inputContract('pptx_apply_format'),
    outputSchema: fixedEditOutput('pptx_apply_format'),
    handler: args => fixedEdit('pptx_apply_format', args),
  },
  {
    name: 'pptx_set_shape_geometry',
    description: 'Set exact native EMU bounds for uniquely identified current-slide PPTX objects. One call batches only this fixed geometry action and does not infer repair coordinates.',
    inputSchema: inputContract('pptx_set_shape_geometry'),
    outputSchema: fixedEditOutput('pptx_set_shape_geometry'),
    handler: args => fixedEdit('pptx_set_shape_geometry', args),
  },
  {
    name: 'pptx_replace_picture_image',
    description: 'Replace embedded PNG or JPEG media for uniquely identified current-slide PPTX pictures while preserving the picture object, geometry, crop, and unrelated media. One call batches only this fixed replacement action.',
    inputSchema: inputContract('pptx_replace_picture_image'),
    outputSchema: fixedEditOutput('pptx_replace_picture_image'),
    handler: args => fixedEdit('pptx_replace_picture_image', args),
  },
  {
    name: 'pptx_validate',
    description: 'Validate a current PPTX package against the published OpenXML contract.',
    inputSchema: inputContract('pptx_validate'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: pptxValidate,
  },
];

function buildServer() {
  const server = new McpServer(
    { name: 'tiwater-office', version: packageMetadata.version },
    {
      instructions: 'Use these tools only for generic Office observation, conversion, editing, validation, and native rendering. Callers own all selected objects, values, and business decisions.',
    },
  );
  for (const tool of tools) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: tool.inputSchema,
        ...(tool.outputSchema ? { outputSchema: tool.outputSchema } : {}),
        ...(tool.annotations ? { annotations: tool.annotations } : {}),
      },
      async args => {
        const payload = await tool.handler(args);
        return createToolResult(payload, { isError: payload?.summary?.pass === false });
      },
    );
  }
  return server;
}

async function docxInspect(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(docxCandidates, ['inspect', input, '--json']);
  return {
    tool: 'docx_inspect',
    runtime: commandRuntime(result),
    source: await fileArtifact(input),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function docxInspectTables(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(docxCandidates, ['inspect-tables', input, '--json']);
  return { tool: 'docx_inspect_tables', runtime: commandRuntime(result), source: await fileArtifact(input), artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json) };
}

async function docxObservation(tool, args) {
  return withTempJsonFile(args, async requestPath => {
    const result = await runJsonCandidateChain(docxCandidates, [tool, requestPath]);
    return { ...result.json, runtime: commandRuntime(result) };
  });
}

async function docxReadObject(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = requireString(args.output, 'output');
  const { output: _output, ...request } = args;
  return withTempJsonFile(request, async requestPath => {
    const result = await runJsonCandidateChain(docxCandidates, ['docx_read_object', requestPath]);
    return {
      tool: 'docx_read_object',
      runtime: commandRuntime(result),
      source: await fileArtifact(input),
      artifact: await writeJsonArtifact(output, result.json),
    };
  });
}

async function docxValidate(args) {
  const result = await runJsonCandidateChain(docxCandidates, ['validate-openxml', requireString(args.input, 'input')], { allowedExitCodes: [0, 1] });
  return { tool: 'docx_validate', runtime: commandRuntime(result), result: result.json };
}

async function docxValidateFontPolicy(args) {
  return withTempJsonFile(args.policy, async policyPath => {
    const result = await runJsonCandidateChain(docxCandidates, ['validate-font-policy', requireString(args.input, 'input'), policyPath], { allowedExitCodes: [0, 1] });
    return { tool: 'docx_validate_font_policy', runtime: commandRuntime(result), result: result.json };
  });
}

async function docxValidateTocStylePolicy(args) {
  const result = await runJsonCandidateChain(docxCandidates, [
    'validate-toc-style-policy', requireString(args.input, 'input'), String(args.italic), String(args.indentCharactersPerLevel),
  ], { allowedExitCodes: [0, 1] });
  return { tool: 'docx_validate_toc_style_policy', runtime: commandRuntime(result), result: result.json };
}

async function docxStripDirectFormatting(args) {
  return copyTransform('docx_strip_direct_formatting', docxCandidates, ['strip-direct-formatting'], args);
}

async function docxReplaceStyleIds(args) {
  return withTempJsonFile(args.styleMap, styleMapPath => copyTransform('docx_replace_style_ids', docxCandidates, ['replace-style-ids'], args, [styleMapPath]));
}

async function copyTransform(tool, candidates, command, args, suffix = []) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  await requireNewFile(output, 'output');
  const result = await runCandidateChain(candidates, [...command, input, output, ...suffix]);
  return { tool, runtime: commandRuntime(result), output: await fileArtifact(output) };
}

async function fixedEdit(tool, args) {
  const output = path.resolve(requireString(args.output, 'output'));
  const receiptOutput = path.resolve(requireString(args.receiptOutput, 'receiptOutput'));
  await requireNewFile(output, 'output');
  await requireNewFile(receiptOutput, 'receiptOutput');
  const candidates = tool.startsWith('docx_') ? docxCandidates
    : tool.startsWith('xlsx_') ? xlsxCandidates
    : pptxCandidates;
  return withTempJsonFile(args, async requestPath => {
    const result = await runJsonCandidateChain(candidates, [tool, requestPath], { allowedExitCodes: [0, 1] });
    if (result.code !== 0) {
      const detail = result.stderr.trim() || result.stdout.trim() || `${tool} failed with exit code ${result.code}`;
      throw new Error(detail);
    }
    if (result.json?.tool !== tool) throw new Error(`${tool} returned a mismatched tool identity`);
    await requireReturnedArtifact(result.json.receipt, receiptOutput, 'receipt');
    if (result.json.output === null) {
      if (result.json.summary?.pass !== false) throw new Error(`${tool} omitted output without reporting failure`);
    } else {
      await requireReturnedArtifact(result.json.output, output, 'output');
      if (result.json.summary?.pass !== true) throw new Error(`${tool} returned output without reporting success`);
    }
    return { ...result.json, runtime: commandRuntime(result) };
  });
}

async function docxCompare(args) {
  const baseline = requireString(args.baseline, 'baseline');
  const updated = requireString(args.updated, 'updated');
  const result = await runJsonCandidateChain(docxCandidates, ['compare', baseline, updated, '--json']);
  return { tool: 'docx_compare', runtime: commandRuntime(result), report: result.json };
}

async function docxExportJson(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(docxCandidates, ['export-json', input]);
  return {
    tool: 'docx_export_json',
    runtime: commandRuntime(result),
    source: await fileArtifact(input),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function officeRenderPdf(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  const receiptOutput = path.resolve(requireString(args.receiptOutput, 'receiptOutput'));
  const sourceFormat = path.extname(input).slice(1).toLowerCase();
  const backend = nativeRenderBackend(sourceFormat);
  if (path.extname(output).toLowerCase() !== '.pdf') {
    throw Object.assign(new Error(`Office render output must be a PDF: ${output}`), { code: -32602 });
  }
  const inputArtifact = await fileArtifact(input);
  await requireNewFile(output, 'output');
  await requireNewFile(receiptOutput, 'receiptOutput');
  await mkdir(path.dirname(output), { recursive: true });
  try {
    const result = await runJsonCandidateChain(
      convertCandidates,
      [`${sourceFormat}-to-pdf`, input, output],
      { env: { TIWATER_OFFICE_PDF_BACKEND: backend } });
    const receipt = nativeRenderReceipt.parse(result.json);
    const pdf = await fileArtifact(output);
    if (path.resolve(receipt.input) !== input
        || path.resolve(receipt.output) !== output
        || receipt.source_format !== sourceFormat
        || receipt.backend !== backend
        || receipt.native_render_provenance.backend !== backend
        || receipt.page_count !== receipt.native_render_provenance.page_count
        || receipt.input_sha256 !== inputArtifact.sha256
        || receipt.native_render_provenance.input.sha256 !== inputArtifact.sha256
        || receipt.native_render_provenance.input.size_bytes !== inputArtifact.bytes
        || receipt.output_sha256 !== pdf.sha256
        || receipt.native_render_provenance.output.sha256 !== pdf.sha256
        || receipt.native_render_provenance.output.size_bytes !== pdf.bytes) {
      throw new Error('Native Office render receipt is not bound to the current input and output');
    }
    return {
      tool: 'office_render_pdf',
      runtime: commandRuntime(result),
      pdf,
      receipt: await writeJsonArtifact(receiptOutput, receipt),
      summary: {
        sourceFormat,
        backend,
        pageCount: receipt.page_count,
      },
    };
  } catch (error) {
    await rm(output, { force: true });
    throw error;
  }
}

async function xlsxConvertLegacy(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  const receiptOutput = path.resolve(requireString(args.receiptOutput, 'receiptOutput'));
  if (path.extname(input).toLowerCase() !== '.xls' || path.extname(output).toLowerCase() !== '.xlsx') {
    throw Object.assign(new Error('Legacy conversion requires .xls input and .xlsx output'), { code: -32602 });
  }
  await requireNewFile(output, 'output');
  await requireNewFile(receiptOutput, 'receiptOutput');
  const inputArtifact = await fileArtifact(input);
  try {
    const result = await runJsonCandidateChain(convertCandidates, ['xls-to-xlsx', input, output]);
    const outputArtifact = await fileArtifact(output);
    const receipt = {
      schema: 'tiwater.office.xlsx-legacy-conversion-receipt/v1',
      backend: 'et', input: inputArtifact, output: outputArtifact, provider: result.json,
    };
    return { tool: 'xlsx_convert_legacy', runtime: commandRuntime(result), output: outputArtifact, receipt: await writeJsonArtifact(receiptOutput, receipt) };
  } catch (error) {
    await rm(output, { force: true });
    throw error;
  }
}

function nativeRenderBackend(sourceFormat) {
  if (['doc', 'docx', 'odt', 'rtf'].includes(sourceFormat)) return 'wps';
  if (['xls', 'xlsx', 'ods'].includes(sourceFormat)) return 'et';
  if (['ppt', 'pptx', 'odp'].includes(sourceFormat)) return 'wpp';
  throw Object.assign(new Error(`Unsupported Office render input: .${sourceFormat || '(none)'}`), { code: -32602 });
}

async function xlsxInspect(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(xlsxCandidates, ['inspect', input, '--json']);
  return {
    tool: 'xlsx_inspect',
    runtime: commandRuntime(result),
    source: await fileArtifact(input),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function xlsxExportJson(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const cmdArgs = ['export-json', input];
  if (args.resolveMergedCells) {
    cmdArgs.push('--resolve-merged-cells');
  }
  const result = await runJsonCandidateChain(xlsxCandidates, cmdArgs);
  return {
    tool: 'xlsx_export_json',
    runtime: commandRuntime(result),
    source: await fileArtifact(input),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function xlsxValidate(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(xlsxCandidates, ['validate', input], { allowedExitCodes: [0, 1] });
  return { tool: 'xlsx_validate', runtime: commandRuntime(result), result: result.json };
}

async function pptxInspect(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(pptxCandidates, ['inspect', input, '--json']);
  return {
    tool: 'pptx_inspect',
    runtime: commandRuntime(result),
    source: await fileArtifact(input),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function pptxExportJson(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(pptxCandidates, ['export-json', input]);
  return {
    tool: 'pptx_export_json',
    runtime: commandRuntime(result),
    source: await fileArtifact(input),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function pptxValidate(args) {
  const result = await runJsonCandidateChain(pptxCandidates, ['validate', requireString(args.input, 'input')], { allowedExitCodes: [0, 1] });
  return { tool: 'pptx_validate', runtime: commandRuntime(result), result: result.json };
}

async function requireReturnedArtifact(returned, expectedPath, label) {
  if (!returned || path.resolve(returned.path || '') !== expectedPath) {
    throw new Error(`${label} artifact is not bound to the accepted provider call`);
  }
  const current = await fileArtifact(expectedPath);
  if (!isDeepStrictEqual(current, returned)) {
    throw new Error(`${label} artifact identity does not match provider output`);
  }
}

async function writeJsonArtifact(output, payload) {
  const fullPath = path.resolve(output);
  await mkdir(path.dirname(fullPath), { recursive: true });
  const bytes = Buffer.from(`${JSON.stringify(payload, null, 2)}\n`, 'utf8');
  await writeFile(fullPath, bytes, { flag: 'wx' });
  return {
    path: fullPath,
    sha256: createHash('sha256').update(bytes).digest('hex'),
    bytes: bytes.length,
  };
}

async function fileArtifact(filePath) {
  const hash = createHash('sha256');
  for await (const chunk of createReadStream(filePath)) {
    hash.update(chunk);
  }
  const file = await stat(filePath);
  return {
    path: path.resolve(filePath),
    sha256: hash.digest('hex'),
    bytes: file.size,
  };
}

async function requireNewFile(filePath, label) {
  try {
    await stat(filePath);
  } catch (error) {
    if (error?.code === 'ENOENT') return;
    throw error;
  }
  throw Object.assign(new Error(`${label} already exists: ${filePath}`), { code: -32602 });
}

function commandRuntime(result) {
  return {
    command: result.command,
    cwd: result.cwd || path.dirname(result.command),
  };
}

serveStdio(buildServer);
