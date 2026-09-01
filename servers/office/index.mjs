#!/usr/bin/env node
import { randomUUID } from 'node:crypto';
import { mkdir, readFile, rename, rm, stat } from 'node:fs/promises';
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
import {
  deliverLargeJsonResult,
  fileArtifact,
  resultChannels,
  returnedContentBudgetBytes,
  writeIdempotentJsonArtifact,
  writeJsonArtifact,
} from '../_shared/large-json-result.mjs';
import { withOutputWriteLock } from '../_shared/output-write-lock.mjs';

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

function largeResultOutput(tool) {
  return z.object({
    tool: z.literal(tool),
    runtime: runtimeIdentity,
    sources: z.array(artifact).min(1).max(2),
    returnContent: z.boolean(),
    artifact: artifact.nullable(),
    receipt: z.object({
      contentBytes: z.number().int().nonnegative(),
      contentReturned: z.boolean(),
      contentWritten: z.boolean(),
    }).strict(),
    content: z.unknown().optional(),
  }).strict();
}

function largeInspectionOutput(tool, summary) {
  return largeResultOutput(tool).extend({ summary }).strict();
}

const xlsxInspectionSummary = z.object({
  sheetCount: z.number().int().nonnegative(),
  sheets: z.array(z.object({
    name: z.string(),
    rowCount: z.number().int().nonnegative(),
    columnCount: z.number().int().nonnegative(),
    usedRange: z.string().nullable(),
    mergedRangeCount: z.number().int().nonnegative(),
    formulaCellCount: z.number().int().nonnegative(),
    openingText: z.array(z.object({
      reference: z.string(),
      textPreview: z.string(),
    }).strict()).max(6),
  }).strict()).max(6),
}).strict();

const pptxInspectionSummary = z.object({
  slideCount: z.number().int().nonnegative(),
  masterCount: z.number().int().nonnegative(),
  slideSize: z.object({
    cx: z.number().int().nonnegative(),
    cy: z.number().int().nonnegative(),
  }).strict().nullable(),
  openingSlides: z.array(z.object({
    slideNumber: z.number().int().positive(),
    textPreview: z.string(),
  }).strict()).max(6),
}).strict();

function docxObservationOutput(tool) {
  return z.object({
    tool: z.literal(tool),
    runtime: runtimeIdentity,
    source: artifact,
    returnContent: z.boolean(),
    artifact: artifact.nullable(),
  }).strict();
}

const docxAddress = z.object({
  part: z.string().min(1),
  path: z.string().startsWith('/'),
}).strict();

function docxInspectionOutput(tool) {
  return docxObservationOutput(tool).extend({
    summary: z.object({
      tableCount: z.number().int().nonnegative(),
      openingParagraphs: z.array(z.object({
        address: docxAddress,
        textPreview: z.string().min(1).max(240),
      }).strict()).max(6),
    }).strict().optional(),
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

const docxListObjectsOutput = docxObservationOutput('docx_list_objects').extend({
  schema: z.literal('tiwater.docx-observation-list/v1'),
  receipt: docxObservationReceipt,
  objects: z.array(docxObjectIdentity).optional(),
  runtime: runtimeIdentity,
}).strict();

const docxTableIndexOutput = docxObservationOutput('docx_table_index').extend({
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
  }).strict()).optional(),
}).strict();

const docxTableReadOutput = docxObservationOutput('docx_read_table').extend({
  schema: z.literal('tiwater.docx-table-page/v1'),
  receipt: z.object({
    schema: z.literal('tiwater.docx-table-page-receipt/v1'),
    totalRowCount: z.number().int().nonnegative(),
    returnedRowCount: z.number().int().nonnegative(),
    remaining: z.number().int().nonnegative(),
    nextOffset: z.number().int().nonnegative().optional(),
    detailPageRetained: z.boolean(),
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
  }).strict()).optional(),
}).strict();

const docxReadObjectOutput = docxObservationOutput('docx_read_object').extend({
  receipt: z.object({
    schema: z.literal('tiwater.docx-read-object-receipt/v1'),
    observationCount: z.number().int().positive(),
    returnedCount: z.number().int().nonnegative(),
    responseComplete: z.boolean(),
    narrowingRequired: z.boolean(),
  }).strict(),
  observations: z.array(docxObservationNode).optional(),
}).strict();

const tools = [
  {
    name: 'docx_inspect',
    description: 'Inspect one current DOCX. Set returnContent true to return the compact identity summary. Provide output to retain the complete machine observation and return its artifact receipt. These channels are independent and may be used together. At least one channel is required. Use list and read operations to traverse selected document objects in native structure order.',
    inputSchema: inputContract('docx_inspect'),
    outputSchema: docxInspectionOutput('docx_inspect'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxInspect,
  },
  {
    name: 'docx_list_objects',
    description: 'Page through mixed nearest-child OpenXML objects only when document order or paragraph relationships are required. Set returnContent true to return a byte-bounded object page. Provide output to store the complete requested page and return its artifact receipt. These channels are independent and may be used together; at least one is required. Continue from receipt.nextOffset only when an unreturned object is needed for the current decision. Do not use this tool to locate tables or read a whole document: use docx_table_index to locate tables, then a narrow table or object read.',
    inputSchema: inputContract('docx_list_objects'),
    outputSchema: docxListObjectsOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: args => docxObservation('docx_list_objects', args),
  },
  {
    name: 'docx_table_index',
    description: 'Locate tables in one current DOCX without returning full cell content or deciding table semantics. Set returnContent true to return a bounded page of addresses, shapes, short previews, and nearest non-empty paragraphs. Provide output to store the complete index and return its artifact receipt. These channels are independent and may be used together; at least one is required. Continue from receipt.nextOffset only when an unreturned table is needed for the current decision, then read one selected native address.',
    inputSchema: inputContract('docx_table_index'),
    outputSchema: docxTableIndexOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: args => docxObservation('docx_table_index', args),
  },
  {
    name: 'docx_read_object',
    description: 'Read explicitly selected rows, cells, or paragraphs from one native DOCX. Set returnContent true to return compact requested descendants; if receipt.narrowingRequired is true, request fewer addresses or descendant kinds. Provide output to store the complete selected observation and return its artifact receipt. These channels are independent and may be used together; at least one is required. A selected cell exposes its vertical-merge owner and logical text, so a continue cell keeps its physical identity while resolving the restart cell value. Use docx_read_table for a table range.',
    inputSchema: inputContract('docx_read_object'),
    outputSchema: docxReadObjectOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxReadObject,
  },
  {
    name: 'docx_read_table',
    description: 'Read an explicit row range from exactly one table selected by native OpenXML address; it never builds another whole-table data object. Set returnContent true to return a byte-bounded compact row page. Provide output to store full paragraph and text-node detail for the selected row page and return its artifact receipt. These channels are independent and may be used together; at least one is required. Request only rows needed for the current decision; receipt.remaining is navigation information, not an obligation to read unused rows, and blank template rows need not be paged through. receipt.nextOffset is present only when another row page exists. Each returned row and cell keeps its reusable native address, zero-based logical gridColumnStart, gridSpan, vertical-merge owner, physical text, and logical text. Match columns across rows by gridColumnStart; a tc[n] path or array position is only that row\'s physical cell ordinal and is not a column identity when gridSpan or gridBefore is present. In a vertical merge, restart begins one logical cell and a continue cell is not an independent row value: logicalText resolves the restart cell value while text remains the physical cell value. Use docx_read_object when one exact object needs a narrower descendant view. The provider reports physical structure only; the Agent decides template and business meaning.',
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
    description: 'Replace the whole text content of paragraph or cell objects observed from this exact input DOCX while retaining target formatting, bookmarks, spans, and vertical merges. For a vertically merged logical cell, write its visible text to the restart cell rather than a continue cell. Tabs and line breaks remain native document text controls; targets containing non-text objects are rejected. This sets already-derived text; it does not insert objects, change table structure, copy source formatting, or decide business wording.',
    inputSchema: inputContract('docx_set_text'),
    outputSchema: fixedEditOutput('docx_set_text'),
    handler: args => fixedEdit('docx_set_text', args),
  },
  {
    name: 'docx_set_table_body',
    description: 'Atomically replace one exact current target-table row range while retaining the table, target styles, grid widths, and all rows and surrounding content outside the range. When retaining leading rows reported with repeatHeader=true, existingRows starts after all of them and never at a verticalMerge=continue row. Name every target grid column in native order and choose one current row inside existingRows as the style prototype for each final row. Horizontal spans use contiguous column IDs. Set rowSpan on one logical cell to occupy multiple rows and omit those columns from the covered rows; the provider writes native vertical merge cells. Every other grid column remains explicit. An empty final row array removes the range when another table row remains. The provider commits once and returns structural readback. It does not read a source table, map source to target, derive text, choose business columns, identify headers, or accept non-text cell content.',
    inputSchema: inputContract('docx_set_table_body'),
    outputSchema: fixedEditOutput('docx_set_table_body'),
    handler: args => fixedEdit('docx_set_table_body', args),
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
    description: 'Compare two DOCX files and report package, metric, and style differences. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('docx_compare'),
    outputSchema: largeResultOutput('docx_compare'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxCompare,
  },
  {
    name: 'docx_export_json',
    description: 'Produce a body-only DOCX JSON projection only when a downstream consumer explicitly requires that format. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required. This does not replace bounded list and read operations.',
    inputSchema: inputContract('docx_export_json'),
    outputSchema: largeResultOutput('docx_export_json'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxExportJson,
  },
  {
    name: 'docx_validate',
    description: 'Validate a current DOCX package against the published OpenXML contract. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('docx_validate'),
    outputSchema: largeResultOutput('docx_validate'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxValidate,
  },
  {
    name: 'docx_validate_font_policy',
    description: 'Validate current DOCX text against an explicit font policy. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('docx_validate_font_policy'),
    outputSchema: largeResultOutput('docx_validate_font_policy'),
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
    description: 'Validate current DOCX table-of-contents paragraph styles against an explicit policy. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('docx_validate_toc_style_policy'),
    outputSchema: largeResultOutput('docx_validate_toc_style_policy'),
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
    description: 'Inspect a current XLSX workbook or legacy XLS workbook, including workbook structure, exported values, formulas, styles, merged ranges, and published legacy-format conversion evidence. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('xlsx_inspect'),
    outputSchema: largeInspectionOutput('xlsx_inspect', xlsxInspectionSummary),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: xlsxInspect,
  },
  {
    name: 'xlsx_export_json',
    description: 'Export workbook sheet data from XLSX as structured JSON. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('xlsx_export_json'),
    outputSchema: largeResultOutput('xlsx_export_json'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: xlsxExportJson,
  },
  ...fixedToolDefinitions(xlsxFixedTools),
  {
    name: 'xlsx_validate',
    description: 'Validate an XLSX workbook package and produce OpenXML validation evidence. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('xlsx_validate'),
    outputSchema: largeResultOutput('xlsx_validate'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: xlsxValidate,
  },
  {
    name: 'pptx_inspect',
    description: 'Inspect a PPTX file, including slides, masters, layouts, shapes, transforms, paragraphs, runs, and placeholders. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('pptx_inspect'),
    outputSchema: largeInspectionOutput('pptx_inspect', pptxInspectionSummary),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: pptxInspect,
  },
  {
    name: 'pptx_export_json',
    description: 'Export PPTX slide text, notes, and placeholder hints as structured JSON. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('pptx_export_json'),
    outputSchema: largeResultOutput('pptx_export_json'),
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
    description: 'Validate a current PPTX package against the published OpenXML contract. Set returnContent true to return the complete result when it fits the response limit. Provide output to write the complete result to a new JSON file. The two choices are independent and may be used together; at least one is required.',
    inputSchema: inputContract('pptx_validate'),
    outputSchema: largeResultOutput('pptx_validate'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: pptxValidate,
  },
];

function buildServer() {
  const server = new McpServer(
    { name: 'tiwater-office', version: packageMetadata.version },
    {
      instructions: 'Use these tools only for generic Office observation, conversion, editing, validation, and native rendering. Callers own all selected objects, values, and business decisions. Every native DOCX address belongs only to the exact input file whose observation returned it and must not be reused with another DOCX. A read-only output path is an immutable artifact identity: an identical request may replay it; every different request uses a different path. Every mutation receiptOutput is a new immutable receipt identity for that exact call and must use a path that has never been used, including when the same document object is updated again.',
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
        const payload = typeof args.output === 'string'
          ? await withOutputWriteLock(args.output, () => tool.handler(args))
          : await tool.handler(args);
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

function compactTableIndexEntry(table) {
  const paragraph = value => value === null ? null : {
    address: value.address,
    textPreview: value.textPreview.trim().replace(/\s+/gu, ' ').slice(0, 160),
    textLength: value.textLength,
  };
  return {
    address: table.address,
    parentAddress: table.parentAddress,
    rowCount: table.rowCount,
    columnCount: table.columnCount,
    textPreview: table.textPreview.trim().replace(/\s+/gu, ' ').slice(0, 240),
    textLength: table.textLength,
    precedingParagraph: paragraph(table.precedingParagraph),
    followingParagraph: paragraph(table.followingParagraph),
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
  const delivery = resultChannels(args);
  const result = await runJsonCandidateChain(docxCandidates, ['inspect', input, '--json']);
  return {
    tool: 'docx_inspect',
    runtime: commandRuntime(result),
    source: await fileArtifact(input),
    returnContent: delivery.returnContent,
    artifact: delivery.output === null
      ? null
      : await writeIdempotentJsonArtifact(delivery.output, result.json),
    ...(delivery.returnContent ? compactDocxInspection(result.json) : {}),
  };
}

async function docxObservation(tool, args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const delivery = resultChannels(args);
  const { output: _output, returnContent: _returnContent, ...request } = args;
  return withTempJsonFile(request, async requestPath => {
    const result = await runJsonCandidateChain(docxCandidates, [tool, requestPath]);
    const payload = { ...result.json, runtime: commandRuntime(result) };
    const retained = {
      tool,
      runtime: payload.runtime,
      source: await fileArtifact(input),
      returnContent: delivery.returnContent,
      artifact: null,
    };
    if (tool === 'docx_list_objects') {
      const totalCount = payload.receipt.totalCount;
      const offset = Math.min(args.offset ?? 0, totalCount);
      const withArtifact = delivery.output === null ? retained : {
        ...retained,
        artifact: await writeIdempotentJsonArtifact(delivery.output, result.json),
      };
      if (!delivery.returnContent) {
        return {
          ...withArtifact,
          schema: payload.schema,
          receipt: payload.receipt,
        };
      }
      const objects = [];
      for (const sourceObject of payload.objects.map(compactDocxObjectIdentity)) {
        const candidate = [...objects, sourceObject];
        if (objects.length > 0
            && Buffer.byteLength(JSON.stringify(candidate)) > returnedContentBudgetBytes) break;
        objects.push(sourceObject);
      }
      const nextOffset = offset + objects.length < totalCount ? offset + objects.length : null;
      return {
        ...withArtifact,
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
      const withArtifact = delivery.output === null ? retained : {
        ...retained,
        artifact: await writeIdempotentJsonArtifact(delivery.output, result.json),
      };
      if (!delivery.returnContent) {
        return {
          ...withArtifact,
          schema: payload.schema,
          receipt: {
            schema: 'tiwater.docx-table-index-receipt/v1',
            totalCount,
            returnedCount: totalCount,
            remaining: 0,
            nextOffset: null,
          },
        };
      }
      const offset = Math.min(args.offset ?? 0, totalCount);
      const requestedLimit = args.limit ?? totalCount;
      const tables = [];
      for (const sourceTable of payload.tables.slice(offset, offset + requestedLimit)) {
        const table = compactTableIndexEntry(sourceTable);
        const candidate = [...tables, table];
        if (tables.length > 0
            && Buffer.byteLength(JSON.stringify(candidate)) > returnedContentBudgetBytes) break;
        tables.push(table);
      }
      const nextOffset = offset + tables.length < totalCount ? offset + tables.length : null;
      return {
        ...withArtifact,
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
      const selectedRows = payload.rows.slice(offset, offset + requestedLimit);
      const selectedNextOffset = offset + selectedRows.length < totalRowCount
        ? offset + selectedRows.length
        : null;
      const selectedPageReceipt = {
        schema: 'tiwater.docx-table-page-receipt/v1',
        totalRowCount,
        returnedRowCount: selectedRows.length,
        remaining: totalRowCount - offset - selectedRows.length,
        ...(selectedNextOffset === null ? {} : { nextOffset: selectedNextOffset }),
      };
      const detailPage = {
        schema: 'tiwater.docx-table-detail-page/v1',
        address: payload.address,
        rowCount: payload.rowCount,
        columnCount: payload.columnCount,
        gridColumns: payload.gridColumns,
        precedingParagraph: payload.precedingParagraph,
        followingParagraph: payload.followingParagraph,
        receipt: selectedPageReceipt,
        rows: selectedRows,
      };
      const withArtifact = delivery.output === null ? retained : {
        ...retained,
        artifact: await writeIdempotentJsonArtifact(delivery.output, detailPage),
      };
      if (!delivery.returnContent) {
        return {
          ...withArtifact,
          schema: 'tiwater.docx-table-page/v1',
          address: payload.address,
          rowCount: payload.rowCount,
          columnCount: payload.columnCount,
          gridColumns: payload.gridColumns,
          precedingParagraph: payload.precedingParagraph,
          followingParagraph: payload.followingParagraph,
          receipt: {
            ...selectedPageReceipt,
            detailPageRetained: true,
            narrowingRequired: false,
          },
        };
      }
      const rows = [];
      let narrowingRequired = false;
      for (const sourceRow of selectedRows) {
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
        if (Buffer.byteLength(JSON.stringify(candidate)) > returnedContentBudgetBytes) {
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
        ...(nextOffset === null ? {} : { nextOffset }),
      };
      return {
        ...withArtifact,
        schema: 'tiwater.docx-table-page/v1',
        address: payload.address,
        rowCount: payload.rowCount,
        columnCount: payload.columnCount,
        gridColumns: payload.gridColumns,
        precedingParagraph: payload.precedingParagraph,
        followingParagraph: payload.followingParagraph,
        receipt: {
          ...pageReceipt,
          detailPageRetained: delivery.output !== null,
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
  const delivery = resultChannels(args);
  const { output: _output, returnContent: _returnContent, ...request } = args;
  return withTempJsonFile(request, async requestPath => {
    const result = await runJsonCandidateChain(docxCandidates, ['docx_read_object', requestPath]);
    const observations = result.json.observations.map(compactDocxObservation);
    const responseComplete = delivery.returnContent
      && Buffer.byteLength(JSON.stringify(observations)) <= returnedContentBudgetBytes;
    return {
      tool: 'docx_read_object',
      runtime: commandRuntime(result),
      source: await fileArtifact(input),
      returnContent: delivery.returnContent,
      artifact: delivery.output === null
        ? null
        : await writeIdempotentJsonArtifact(delivery.output, result.json),
      receipt: {
        schema: 'tiwater.docx-read-object-receipt/v1',
        observationCount: observations.length,
        returnedCount: responseComplete ? observations.length : 0,
        responseComplete,
        narrowingRequired: delivery.returnContent && !responseComplete,
      },
      ...(responseComplete ? { observations } : {}),
    };
  });
}

async function docxValidate(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(docxCandidates, ['validate-openxml', input], { allowedExitCodes: [0, 1] });
  return deliverLargeJsonResult({ tool: 'docx_validate', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

async function docxValidateFontPolicy(args) {
  return withTempJsonFile(args.policy, async policyPath => {
    const input = path.resolve(requireString(args.input, 'input'));
    const result = await runJsonCandidateChain(docxCandidates, ['validate-font-policy', input, policyPath], { allowedExitCodes: [0, 1] });
    return deliverLargeJsonResult({ tool: 'docx_validate_font_policy', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
  });
}

async function docxValidateTocStylePolicy(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(docxCandidates, [
    'validate-toc-style-policy', input, String(args.italic), String(args.indentCharactersPerLevel),
  ], { allowedExitCodes: [0, 1] });
  return deliverLargeJsonResult({ tool: 'docx_validate_toc_style_policy', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
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
  const baseline = path.resolve(requireString(args.baseline, 'baseline'));
  const updated = path.resolve(requireString(args.updated, 'updated'));
  const result = await runJsonCandidateChain(docxCandidates, ['compare', baseline, updated, '--json']);
  return deliverLargeJsonResult({ tool: 'docx_compare', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [baseline, updated] });
}

async function docxExportJson(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(docxCandidates, ['export-json', input]);
  return deliverLargeJsonResult({ tool: 'docx_export_json', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
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
  return deliverLargeJsonResult({
    tool: 'xlsx_inspect',
    args,
    runtime: commandRuntime(result),
    payload: result.json,
    sourcePaths: [input],
    summary: compactXlsxInspection(result.json),
  });
}

async function xlsxExportJson(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const cmdArgs = ['export-json', input];
  if (args.resolveMergedCells) {
    cmdArgs.push('--resolve-merged-cells');
  }
  const result = await runJsonCandidateChain(xlsxCandidates, cmdArgs);
  return deliverLargeJsonResult({ tool: 'xlsx_export_json', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

async function xlsxValidate(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(xlsxCandidates, ['validate', input], { allowedExitCodes: [0, 1] });
  return deliverLargeJsonResult({ tool: 'xlsx_validate', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

async function pptxInspect(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(pptxCandidates, ['inspect', input, '--json']);
  return deliverLargeJsonResult({
    tool: 'pptx_inspect',
    args,
    runtime: commandRuntime(result),
    payload: result.json,
    sourcePaths: [input],
    summary: compactPptxInspection(result.json),
  });
}

async function pptxExportJson(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(pptxCandidates, ['export-json', input]);
  return deliverLargeJsonResult({ tool: 'pptx_export_json', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

async function pptxValidate(args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const result = await runJsonCandidateChain(pptxCandidates, ['validate', input], { allowedExitCodes: [0, 1] });
  return deliverLargeJsonResult({ tool: 'pptx_validate', args, runtime: commandRuntime(result), payload: result.json, sourcePaths: [input] });
}

function compactPreview(value, maximum) {
  return typeof value === 'string'
    ? value.trim().replace(/\s+/gu, ' ').slice(0, maximum)
    : '';
}

function compactXlsxInspection(payload) {
  const workbook = payload?.workbook ?? {};
  const allSheets = Array.isArray(workbook.sheets) ? workbook.sheets : [];
  return {
    sheetCount: Number.isInteger(workbook.sheetCount) ? workbook.sheetCount : allSheets.length,
    sheets: allSheets.slice(0, 6).map(sheet => ({
      name: compactPreview(sheet?.name, 120),
      rowCount: Number.isInteger(sheet?.rowCount) ? sheet.rowCount : 0,
      columnCount: Number.isInteger(sheet?.columnCount) ? sheet.columnCount : 0,
      usedRange: typeof sheet?.usedRange === 'string' ? compactPreview(sheet.usedRange, 120) : null,
      mergedRangeCount: Array.isArray(sheet?.mergedRanges) ? sheet.mergedRanges.length : 0,
      formulaCellCount: Number.isInteger(sheet?.formulaCellCount) ? sheet.formulaCellCount : 0,
      openingText: (Array.isArray(sheet?.textCells) ? sheet.textCells : [])
        .filter(cell => typeof cell?.text === 'string' && cell.text.trim() !== '')
        .slice(0, 6)
        .map(cell => ({
          reference: compactPreview(cell.reference, 40),
          textPreview: compactPreview(cell.text, 160),
        })),
    })),
  };
}

function compactPptxInspection(payload) {
  const slides = Array.isArray(payload?.slides) ? payload.slides : [];
  return {
    slideCount: Number.isInteger(payload?.slideCount) ? payload.slideCount : slides.length,
    masterCount: Array.isArray(payload?.masters) ? payload.masters.length : 0,
    slideSize: Number.isInteger(payload?.slideSize?.cx) && Number.isInteger(payload?.slideSize?.cy)
      ? { cx: payload.slideSize.cx, cy: payload.slideSize.cy }
      : null,
    openingSlides: slides.slice(0, 6).map((slide, index) => ({
      slideNumber: Number.isInteger(slide?.slideNumber) && slide.slideNumber > 0
        ? slide.slideNumber
        : index + 1,
      textPreview: compactPreview(
        (Array.isArray(slide?.shapes) ? slide.shapes : [])
          .map(shape => shape?.text)
          .filter(text => typeof text === 'string' && text.trim() !== '')
          .join(' '),
        240,
      ),
    })),
  };
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
