#!/usr/bin/env node
import { createHash, randomUUID } from 'node:crypto';
import { createReadStream } from 'node:fs';
import { mkdir, readFile, rename, rm, stat, writeFile } from 'node:fs/promises';
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

const docxFieldRefreshReceipt = z.object({
  schema: z.literal('tiwater.convert-refresh-docx-fields/v1'),
  status: z.literal('ok'),
  input: z.string(),
  input_sha256: z.string().regex(/^[a-f0-9]{64}$/),
  output: z.string(),
  output_sha256: z.string().regex(/^[a-f0-9]{64}$/),
  source_format: z.literal('docx'),
  target_format: z.literal('docx'),
  version: z.string(),
  backend: z.literal('wps'),
  refresh_scope: z.array(z.enum(['table-of-contents', 'table-of-figures'])).min(1),
}).strict();

const docxFieldRefreshOutput = z.object({
  tool: z.literal('docx_refresh_fields'),
  runtime: runtimeIdentity,
  output: artifact,
  receipt: artifact,
  summary: z.object({
    backend: z.literal('wps'),
    refreshScope: z.array(z.enum(['table-of-contents', 'table-of-figures'])).min(1),
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

const docxAddress = z.object({
  part: z.string().min(1),
  path: z.string().startsWith('/'),
}).strict();

function docxInspectionOutput(tool) {
  return artifactOutput(tool).extend({
    summary: z.object({
      tableCount: z.number().int().nonnegative(),
      openingParagraphs: z.array(z.object({
        address: docxAddress,
        textPreview: z.string().min(1).max(240),
      }).strict()).max(6),
    }).strict(),
  }).strict();
}

const docxObjectIdentity = z.object({
  address: docxAddress,
  parentAddress: docxAddress.nullable(),
  kind: z.string(),
  textPreview: z.string().nullable(),
  gridSpan: z.number().int().positive().nullable(),
  verticalMerge: z.string().nullable(),
}).strict();

const docxNestedObjectIdentity = docxObjectIdentity.pick({
  address: true,
  kind: true,
  textPreview: true,
}).extend({
  gridSpan: z.number().int().positive().optional(),
  verticalMerge: z.string().optional(),
  verticalMergeOwner: docxAddress.optional(),
  logicalText: z.string().optional(),
}).strict();
const docxObservationNode = z.lazy(() => z.object({
  object: docxNestedObjectIdentity,
  children: z.array(docxObservationNode).optional(),
}).strict());

const docxObservationReceipt = z.object({
  schema: z.literal('tiwater.docx-observation-receipt/v1'),
  operation: z.literal('list'),
  totalCount: z.number().int().nonnegative(),
  returnedCount: z.number().int().nonnegative(),
  remaining: z.number().int().nonnegative(),
  nextOffset: z.number().int().nonnegative().nullable(),
}).strict();

const docxListObjectsOutput = artifactOutput('docx_list_objects').extend({
  schema: z.literal('tiwater.docx-observation-list/v1'),
  receipt: docxObservationReceipt,
  objects: z.array(docxObjectIdentity),
  runtime: runtimeIdentity,
}).strict();

const docxTableIndexOutput = artifactOutput('docx_table_index').extend({
  schema: z.literal('tiwater.docx-table-index/v1'),
  receipt: z.object({
    schema: z.literal('tiwater.docx-table-index-receipt/v1'),
    totalCount: z.number().int().nonnegative(),
    returnedCount: z.number().int().nonnegative(),
    remaining: z.number().int().nonnegative(),
    nextOffset: z.number().int().nonnegative().nullable(),
  }).strict(),
  tables: z.array(z.object({
    address: docxAddress,
    parentAddress: docxAddress.nullable(),
    rowCount: z.number().int().nonnegative(),
    columnCount: z.number().int().nonnegative(),
    textPreview: z.string(),
    textLength: z.number().int().nonnegative(),
    precedingParagraph: z.object({
      address: docxAddress,
      textPreview: z.string(),
      textLength: z.number().int().nonnegative(),
    }).strict().nullable(),
    followingParagraph: z.object({
      address: docxAddress,
      textPreview: z.string(),
      textLength: z.number().int().nonnegative(),
    }).strict().nullable(),
  }).strict()),
}).strict();

const docxTableReadOutput = artifactOutput('docx_read_table').extend({
  schema: z.literal('tiwater.docx-table-page/v1'),
  receipt: z.object({
    schema: z.literal('tiwater.docx-table-page-receipt/v1'),
    totalRowCount: z.number().int().nonnegative(),
    returnedRowCount: z.number().int().nonnegative(),
    remaining: z.number().int().nonnegative(),
    nextOffset: z.number().int().nonnegative().nullable(),
    detailPageRetained: z.literal(true),
    narrowingRequired: z.boolean(),
  }).strict(),
  address: docxAddress,
  rowCount: z.number().int().nonnegative(),
  columnCount: z.number().int().nonnegative(),
  gridColumns: z.array(z.object({
    address: docxAddress,
    widthTwips: z.number().int().positive().nullable(),
  }).strict()),
  precedingParagraph: z.object({
    address: docxAddress,
    textPreview: z.string(),
    textLength: z.number().int().nonnegative(),
  }).strict().nullable(),
  followingParagraph: z.object({
    address: docxAddress,
    textPreview: z.string(),
    textLength: z.number().int().nonnegative(),
  }).strict().nullable(),
  rows: z.array(z.object({
    address: docxAddress,
    repeatHeader: z.boolean(),
    cantSplit: z.boolean(),
    gridBefore: z.number().int().nonnegative(),
    gridAfter: z.number().int().nonnegative(),
    cells: z.array(z.object({
      address: docxAddress,
      gridColumnStart: z.number().int().nonnegative(),
      gridSpan: z.number().int().positive(),
      verticalMerge: z.string().nullable(),
      verticalMergeOwner: docxAddress.nullable(),
      text: z.string(),
      logicalText: z.string(),
    }).strict()),
  }).strict()),
}).strict();

const docxReadObjectOutput = artifactOutput('docx_read_object').extend({
  receipt: z.object({
    schema: z.literal('tiwater.docx-read-object-receipt/v1'),
    observationCount: z.number().int().positive(),
    returnedCount: z.number().int().nonnegative(),
    responseComplete: z.boolean(),
    narrowingRequired: z.boolean(),
  }).strict(),
  observations: z.array(docxObservationNode),
}).strict();

const tools = [
  {
    name: 'docx_inspect',
    description: 'Inspect one current DOCX and retain its complete machine observation at output. The response returns the observation artifact, table count, and up to six opening non-empty body paragraphs with their OpenXML addresses. Use list and read operations to traverse selected document objects in native structure order.',
    inputSchema: inputContract('docx_inspect'),
    outputSchema: docxInspectionOutput('docx_inspect'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxInspect,
  },
  {
    name: 'docx_list_objects',
    description: 'Page through mixed nearest-child OpenXML objects when document order or paragraph relationships are required. The complete requested provider page is retained at output; the response returns a byte-bounded prefix. Continue with receipt.nextOffset and descend through a returned parent address. Do not use this tool to locate tables or read a whole document: use docx_table_index to locate tables, then docx_read_table for one selected table.',
    inputSchema: inputContract('docx_list_objects'),
    outputSchema: docxListObjectsOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: args => docxObservation('docx_list_objects', args),
  },
  {
    name: 'docx_table_index',
    description: 'Locate tables in one current DOCX. Writes the complete index to output and returns a bounded page containing table address, shape, short text preview, and nearest non-empty paragraphs. Continue with receipt.nextOffset, then call docx_read_table for one selected address. It does not return full cell content or decide table semantics.',
    inputSchema: inputContract('docx_table_index'),
    outputSchema: docxTableIndexOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: args => docxObservation('docx_table_index', args),
  },
  {
    name: 'docx_read_object',
    description: 'Read selected rows, cells, or paragraphs from one native DOCX and retain the complete result at output. A selected cell exposes its vertical-merge owner and logical text, so a continue cell keeps its physical identity while resolving the restart cell value. The response returns compact requested descendants when bounded; if receipt.narrowingRequired is true, request fewer observed addresses or descendants. Use docx_read_table for a table.',
    inputSchema: inputContract('docx_read_object'),
    outputSchema: docxReadObjectOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxReadObject,
  },
  {
    name: 'docx_read_table',
    description: 'Read exactly one table selected from docx_table_index by native OpenXML address. Each call retains full paragraph and text-node detail for exactly the returned row page at output; it never builds another whole-table data object. The machine response is the compact form of that page: each row and cell keeps its reusable native address, grid position and span, vertical-merge owner, physical text, and logical text. In a vertical merge, restart begins one logical cell and a continue cell is not an independent row value: logicalText resolves the restart cell value while text remains the physical cell value. A table observation is complete only when receipt.remaining is 0; process each page once in ascending receipt.nextOffset order. Use docx_read_object only when one selected object needs a narrower descendant view. The provider reports physical structure only; the Agent decides the template and business meaning.',
    inputSchema: inputContract('docx_read_table'),
    outputSchema: docxTableReadOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: args => docxObservation('docx_read_table', args),
  },
  {
    name: 'docx_copy_content',
    description: 'Replace content in existing plain-text target paragraphs or cells while retaining target container formatting. A whole-object selection copies only that object\'s inline content; a range copies an exact substring from a run or text address. Source table, row, cell, span, and merge structure are not copied.',
    inputSchema: inputContract('docx_copy_content'),
    outputSchema: fixedEditOutput('docx_copy_content'),
    handler: args => fixedEdit('docx_copy_content', args),
  },
  {
    name: 'docx_set_text',
    description: 'Replace the whole text content of observed paragraph or cell objects while retaining target formatting, bookmarks, spans, and vertical merges. For a vertically merged logical cell, write its visible text to the restart cell rather than a continue cell. Tabs and line breaks remain native document text controls; targets containing non-text objects are rejected. This sets already-derived text; it does not insert objects, change table structure, copy source formatting, or decide business wording.',
    inputSchema: inputContract('docx_set_text'),
    outputSchema: fixedEditOutput('docx_set_text'),
    handler: args => fixedEdit('docx_set_text', args),
  },
  {
    name: 'docx_insert_objects',
    description: 'Insert selected current DOCX objects under an existing parent. Table rows are objects: expand a target table by copying one contiguous observed row range and use repeat for count; sourceInput may equal input. A row range beginning with vertical-merge continuations may be inserted only inside a target boundary with the same active grid spans, which extends those merges. Individual table cells are not raw insertion targets because that would bypass the table grid.',
    inputSchema: inputContract('docx_insert_objects'),
    outputSchema: fixedEditOutput('docx_insert_objects'),
    handler: args => fixedEdit('docx_insert_objects', args),
  },
  {
    name: 'docx_delete_object',
    description: 'Delete selected current DOCX objects directly from the current target document. Selected table rows must close every vertical merge and cannot remove the whole table. Individual table cells are not raw deletion targets; use column or merge operations for table structure.',
    inputSchema: inputContract('docx_delete_object'),
    outputSchema: fixedEditOutput('docx_delete_object'),
    handler: args => fixedEdit('docx_delete_object', args),
  },
  {
    name: 'docx_merge_cells',
    description: 'Merge selected current DOCX cells when they form one closed rectangle. A one-column, multi-row rectangle creates a vertical merge whose first cell is the restart owner and whose later cells are continuations. All selected cell content moves into the top-left owner, so the selected content must already be correct for that one logical cell.',
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
    name: 'docx_insert_table_columns',
    description: 'Insert empty template-shaped grid columns into one current main-document table. Select an observed source grid column for width and per-row cell formatting, and optionally a before grid-column address; cells spanning the insertion boundary expand instead of being split. It does not copy business values or decide column meaning.',
    inputSchema: inputContract('docx_insert_table_columns'),
    outputSchema: fixedEditOutput('docx_insert_table_columns'),
    handler: args => fixedEdit('docx_insert_table_columns', args),
  },
  {
    name: 'docx_delete_table_columns',
    description: 'Delete selected observed grid columns from one current main-document table while shrinking spanning cells and preserving the remaining table grid. It cannot remove every column and does not decide whether a business column is unused.',
    inputSchema: inputContract('docx_delete_table_columns'),
    outputSchema: fixedEditOutput('docx_delete_table_columns'),
    handler: args => fixedEdit('docx_delete_table_columns', args),
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
    annotations: { readOnlyHint: true, idempotentHint: true },
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
    name: 'docx_apply_font_policy',
    description: 'Apply one explicit font family and size policy to current main-document body and table text. It does not derive a policy or alter other run semantics.',
    inputSchema: inputContract('docx_apply_font_policy'),
    outputSchema: fixedEditOutput('docx_apply_font_policy'),
    handler: args => fixedEdit('docx_apply_font_policy', args),
  },
  {
    name: 'docx_validate_toc_style_policy',
    description: 'Validate current DOCX table-of-contents paragraph styles against an explicit policy.',
    inputSchema: inputContract('docx_validate_toc_style_policy'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxValidateTocStylePolicy,
  },
  {
    name: 'docx_apply_toc_style_policy',
    description: 'Apply explicit italic and per-level indentation values to current built-in table-of-contents paragraph styles. It does not change heading text or refresh fields.',
    inputSchema: inputContract('docx_apply_toc_style_policy'),
    outputSchema: fixedEditOutput('docx_apply_toc_style_policy'),
    handler: args => fixedEdit('docx_apply_toc_style_policy', args),
  },
  {
    name: 'docx_refresh_fields',
    description: 'Refresh table-of-contents and table-of-figures field results in a current DOCX through native WPS Writer. It does not change headings, captions, or field definitions.',
    inputSchema: inputContract('docx_refresh_fields'),
    outputSchema: docxFieldRefreshOutput,
    handler: docxRefreshFields,
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
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: xlsxInspect,
  },
  {
    name: 'xlsx_export_json',
    description: 'Export workbook sheet data from XLSX as structured JSON.',
    inputSchema: inputContract('xlsx_export_json'),
    outputSchema: artifactOutput('xlsx_export_json'),
    annotations: { readOnlyHint: true, idempotentHint: true },
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
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: pptxInspect,
  },
  {
    name: 'pptx_export_json',
    description: 'Export PPTX slide text, notes, and placeholder hints to a new JSON artifact without returning the full presentation through MCP.',
    inputSchema: inputContract('pptx_export_json'),
    outputSchema: artifactOutput('pptx_export_json'),
    annotations: { readOnlyHint: true, idempotentHint: true },
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

function compactDocxInspection(report) {
  if (!Array.isArray(report?.tables?.tables) || !Array.isArray(report.flow)) {
    throw new Error('docx-table-inspection-result-invalid');
  }
  const openingParagraphs = report.flow
    .filter(item => item?.type === 'paragraph'
      && typeof item.address?.part === 'string'
      && typeof item.address?.path === 'string'
      && typeof item.text === 'string' && item.text.trim())
    .slice(0, 6)
    .map(item => ({
      address: item.address,
      textPreview: item.text.trim().replace(/\s+/gu, ' ').slice(0, 240),
    }));
  return {
    summary: {
      tableCount: report.tables.tables.length,
      openingParagraphs,
    },
  };
}

function compactDocxObjectIdentity(object) {
  return {
    address: object.address,
    parentAddress: object.parentAddress,
    kind: object.kind,
    textPreview: object.textPreview,
    gridSpan: object.gridSpan,
    verticalMerge: object.verticalMerge,
  };
}

function compactDocxObservation(observation) {
  if (!observation?.object) throw new Error('docx-read-object-missing-observation');
  function compactNode(node) {
    const identity = compactDocxObjectIdentity(node.object);
    const children = (node.children ?? []).map(compactNode);
    const object = {
      address: identity.address,
      kind: identity.kind,
      textPreview: children.length === 0 ? identity.textPreview : null,
      ...(identity.gridSpan === null ? {} : { gridSpan: identity.gridSpan }),
      ...(identity.verticalMerge === null ? {} : { verticalMerge: identity.verticalMerge }),
      ...(node.object.verticalMergeOwner === null ? {} : { verticalMergeOwner: node.object.verticalMergeOwner }),
      ...(node.object.logicalText === null ? {} : { logicalText: node.object.logicalText }),
    };
    return {
      object,
      ...(children.length === 0 ? {} : { children }),
    };
  }
  return compactNode(observation);
}

async function docxInspect(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(docxCandidates, ['inspect', input, '--json']);
  return {
    tool: 'docx_inspect',
    runtime: commandRuntime(result),
    source: await fileArtifact(input),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
    ...compactDocxInspection(result.json),
  };
}

async function docxObservation(tool, args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = requireString(args.output, 'output');
  const { output: _output, ...request } = args;
  return withTempJsonFile(request, async requestPath => {
    const result = await runJsonCandidateChain(docxCandidates, [tool, requestPath]);
    const payload = { ...result.json, runtime: commandRuntime(result) };
    const retained = {
      tool,
      runtime: payload.runtime,
      source: await fileArtifact(input),
      ...(tool === 'docx_read_table' ? {} : {
        artifact: await writeJsonArtifact(output, result.json),
      }),
    };
    if (tool === 'docx_list_objects') {
      const totalCount = payload.receipt.totalCount;
      const offset = Math.min(args.offset ?? 0, totalCount);
      const objects = [];
      for (const sourceObject of payload.objects.map(compactDocxObjectIdentity)) {
        const candidate = [...objects, sourceObject];
        if (objects.length > 0 && Buffer.byteLength(JSON.stringify(candidate)) > 6_500) break;
        objects.push(sourceObject);
      }
      const nextOffset = offset + objects.length < totalCount ? offset + objects.length : null;
      return {
        ...retained,
        schema: payload.schema,
        receipt: {
          schema: 'tiwater.docx-observation-receipt/v1',
          operation: 'list',
          totalCount,
          returnedCount: objects.length,
          remaining: totalCount - offset - objects.length,
          nextOffset,
        },
        objects,
      };
    }
    if (tool === 'docx_table_index') {
      const totalCount = payload.tables.length;
      const offset = Math.min(args.offset ?? 0, totalCount);
      const requestedLimit = args.limit ?? totalCount;
      const tables = [];
      for (const table of payload.tables.slice(offset, offset + requestedLimit)) {
        const candidate = [...tables, table];
        if (tables.length > 0 && Buffer.byteLength(JSON.stringify(candidate)) > 7_000) break;
        tables.push(table);
      }
      const nextOffset = offset + tables.length < totalCount ? offset + tables.length : null;
      return {
        ...retained,
        schema: payload.schema,
        receipt: {
          schema: 'tiwater.docx-table-index-receipt/v1',
          totalCount,
          returnedCount: tables.length,
          remaining: totalCount - offset - tables.length,
          nextOffset,
        },
        tables,
      };
    }
    if (tool === 'docx_read_table') {
      const totalRowCount = payload.rows.length;
      const offset = Math.min(args.offset ?? 0, totalRowCount);
      const requestedLimit = args.limit ?? totalRowCount;
      const rows = [];
      let narrowingRequired = false;
      for (const sourceRow of payload.rows.slice(offset, offset + requestedLimit)) {
        const row = {
          address: sourceRow.address,
          repeatHeader: sourceRow.repeatHeader,
          cantSplit: sourceRow.cantSplit,
          gridBefore: sourceRow.gridBefore,
          gridAfter: sourceRow.gridAfter,
          cells: sourceRow.cells.map(cell => ({
            address: cell.address,
            gridColumnStart: cell.gridColumnStart,
            gridSpan: cell.gridSpan,
            verticalMerge: cell.verticalMerge,
            verticalMergeOwner: cell.verticalMergeOwner,
            text: cell.paragraphs.map(paragraph => paragraph.text).join('\n'),
            logicalText: cell.logicalText,
          })),
        };
        const candidate = [...rows, row];
        if (Buffer.byteLength(JSON.stringify(candidate)) > 6_500) {
          narrowingRequired = rows.length === 0;
          break;
        }
        rows.push(row);
      }
      const nextOffset = !narrowingRequired && offset + rows.length < totalRowCount
        ? offset + rows.length
        : null;
      const pageReceipt = {
        schema: 'tiwater.docx-table-page-receipt/v1',
        totalRowCount,
        returnedRowCount: rows.length,
        remaining: totalRowCount - offset - rows.length,
        nextOffset,
      };
      const detailPage = {
        schema: 'tiwater.docx-table-detail-page/v1',
        address: payload.address,
        rowCount: payload.rowCount,
        columnCount: payload.columnCount,
        gridColumns: payload.gridColumns,
        precedingParagraph: payload.precedingParagraph,
        followingParagraph: payload.followingParagraph,
        receipt: pageReceipt,
        rows: payload.rows.slice(offset, offset + rows.length),
      };
      return {
        ...retained,
        artifact: await writeJsonArtifact(output, detailPage),
        schema: 'tiwater.docx-table-page/v1',
        address: payload.address,
        rowCount: payload.rowCount,
        columnCount: payload.columnCount,
        gridColumns: payload.gridColumns,
        precedingParagraph: payload.precedingParagraph,
        followingParagraph: payload.followingParagraph,
        receipt: {
          ...pageReceipt,
          detailPageRetained: true,
          narrowingRequired,
        },
        rows,
      };
    }
    throw new Error(`unsupported-docx-observation-tool:${tool}`);
  });
}

async function docxReadObject(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = requireString(args.output, 'output');
  const { output: _output, ...request } = args;
  return withTempJsonFile(request, async requestPath => {
    const result = await runJsonCandidateChain(docxCandidates, ['docx_read_object', requestPath]);
    const observations = result.json.observations.map(compactDocxObservation);
    const responseComplete = Buffer.byteLength(JSON.stringify(observations)) <= 6_500;
    return {
      tool: 'docx_read_object',
      runtime: commandRuntime(result),
      source: await fileArtifact(input),
      artifact: await writeJsonArtifact(output, result.json),
      receipt: {
        schema: 'tiwater.docx-read-object-receipt/v1',
        observationCount: observations.length,
        returnedCount: responseComplete ? observations.length : 0,
        responseComplete,
        narrowingRequired: !responseComplete,
      },
      observations: responseComplete ? observations : [],
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

async function docxRefreshFields(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  const receiptOutput = path.resolve(requireString(args.receiptOutput, 'receiptOutput'));
  const inPlace = input === output;
  const invocationOutput = inPlace ? temporarySibling(output) : output;
  if (path.extname(input).toLowerCase() !== '.docx' || path.extname(output).toLowerCase() !== '.docx') {
    throw Object.assign(new Error('DOCX field refresh requires .docx input and output'), { code: -32602 });
  }
  if (!inPlace) await requireNewFile(output, 'output');
  await requireNewFile(receiptOutput, 'receiptOutput');
  const inputArtifact = await fileArtifact(input);
  try {
    const result = await runJsonCandidateChain(convertCandidates, ['refresh-docx-fields', input, invocationOutput]);
    const providerReceipt = docxFieldRefreshReceipt.parse(result.json);
    const invocationArtifact = await fileArtifact(invocationOutput);
    if (path.resolve(providerReceipt.input) !== input
        || path.resolve(providerReceipt.output) !== invocationOutput
        || providerReceipt.input_sha256 !== inputArtifact.sha256
        || providerReceipt.output_sha256 !== invocationArtifact.sha256) {
      throw new Error('DOCX field refresh receipt is not bound to the current input and output');
    }
    if (inPlace) await rename(invocationOutput, output);
    const outputArtifact = await fileArtifact(output);
    const acceptedReceipt = { ...providerReceipt, output, output_sha256: outputArtifact.sha256 };
    return {
      tool: 'docx_refresh_fields',
      runtime: commandRuntime(result),
      output: outputArtifact,
      receipt: await writeJsonArtifact(receiptOutput, acceptedReceipt),
      summary: {
        backend: providerReceipt.backend,
        refreshScope: providerReceipt.refresh_scope,
      },
    };
  } catch (error) {
    await rm(invocationOutput, { force: true });
    throw error;
  }
}

async function copyTransform(tool, candidates, command, args, suffix = []) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  const inPlace = input === output;
  const invocationOutput = inPlace ? temporarySibling(output) : output;
  if (!inPlace) await requireNewFile(output, 'output');
  try {
    const result = await runCandidateChain(candidates, [...command, input, invocationOutput, ...suffix]);
    if (inPlace) await rename(invocationOutput, output);
    return { tool, runtime: commandRuntime(result), output: await fileArtifact(output) };
  } catch (error) {
    await rm(invocationOutput, { force: true });
    throw error;
  }
}

async function fixedEdit(tool, args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  const receiptOutput = path.resolve(requireString(args.receiptOutput, 'receiptOutput'));
  if (!(tool.startsWith('docx_') && input === output)) await requireNewFile(output, 'output');
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

function temporarySibling(filePath) {
  const extension = path.extname(filePath);
  const stem = path.basename(filePath, extension);
  return path.join(path.dirname(filePath), `.${stem}.${randomUUID()}.tmp${extension}`);
}

function commandRuntime(result) {
  return {
    command: result.command,
    cwd: result.cwd || path.dirname(result.command),
  };
}

serveStdio(buildServer);
