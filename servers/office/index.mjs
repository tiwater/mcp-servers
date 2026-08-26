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
const invocationCwd = process.cwd();

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

const pathInput = z.string().trim().min(1);

const runtimeIdentity = z.object({
  command: z.string(),
  cwd: z.string(),
}).strict();

const inputOnly = z.object({ input: pathInput }).strict();
const artifactInput = z.object({
  input: pathInput,
  output: pathInput.describe('New JSON artifact path. Existing files are never overwritten.'),
}).strict();
const artifact = z.object({
  path: z.string(),
  sha256: z.string().regex(/^[0-9a-f]{64}$/),
  bytes: z.number().int().nonnegative(),
}).strict();

const pptxTemplateApplyResult = z.object({
  input: z.string().min(1),
  template: z.string().min(1),
  output: z.string().min(1),
  changedSlideCount: z.number().int().nonnegative(),
  issues: z.array(z.object({
    slideNumber: z.number().int().positive().nullable(),
    message: z.string(),
  }).strict()),
  materializedLayoutShapes: z.array(z.object({
    slideNumber: z.number().int().positive(),
    sourceLayoutPath: z.string().min(1),
    sourceShapeId: z.number().int().nonnegative().max(0xffffffff),
    outputShapeId: z.number().int().nonnegative().max(0xffffffff),
  }).strict()),
  frozenPlaceholderCount: z.number().int().nonnegative(),
  removedSystemPlaceholders: z.array(z.object({
    slideNumber: z.number().int().positive(),
    shapeId: z.number().int().nonnegative().max(0xffffffff),
    placeholderType: z.string().min(1),
  }).strict()),
}).strict();
const pptxFormatApplyResult = z.object({
  input: z.string().min(1),
  output: z.string().min(1),
  operationCount: z.number().int().nonnegative(),
  changedCount: z.number().int().nonnegative(),
  changes: z.array(z.object({
    slideNumber: z.number().int().positive(),
    shapeId: z.number().int().nonnegative().max(0xffffffff),
    runIndex: z.number().int().nonnegative(),
    properties: z.array(z.string()),
  }).strict()),
  issues: z.array(z.object({
    slideNumber: z.number().int().positive(),
    shapeId: z.number().int().nonnegative().max(0xffffffff),
    runIndex: z.number().int().nonnegative(),
    message: z.string(),
  }).strict()),
}).strict();
const pptxTemplateApplyOutput = z.object({
  tool: z.literal('pptx_apply_template'),
  runtime: runtimeIdentity,
  receipt: artifact,
  output: artifact.nullable(),
  summary: z.object({
    pass: z.boolean(),
    changedSlideCount: z.number().int().nonnegative(),
    issueCount: z.number().int().nonnegative(),
  }).strict(),
}).strict();
const pptxFormatApplyOutput = z.object({
  tool: z.literal('pptx_apply_format'),
  runtime: runtimeIdentity,
  receipt: artifact,
  output: artifact.nullable(),
  summary: z.object({
    pass: z.boolean(),
    operationCount: z.number().int().nonnegative(),
    changedCount: z.number().int().nonnegative(),
    issueCount: z.number().int().nonnegative(),
  }).strict(),
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

const index = z.number().int().nonnegative();
const positiveIndex = z.number().int().positive();
const optionalTextMatch = {
  matchMode: z.string().optional(),
  paragraphStyle: z.string().optional(),
};
const richTextSegment = z.object({
  text: z.string(),
  color: z.string().optional(),
  underline: z.boolean().optional(),
  bold: z.boolean().optional(),
  fontName: z.string().optional(),
}).strict();
const tableCellInput = z.object({
  text: z.string().optional(),
  gridSpan: positiveIndex.optional(),
  vMerge: z.string().optional(),
  bold: z.boolean().optional(),
  header: z.boolean().optional(),
  shading: z.string().optional(),
  alignment: z.string().optional(),
  richText: z.array(richTextSegment).optional(),
}).strict();

function editAction(name, operationType, description, changeSchema) {
  return { name, operationType, description, changeSchema, batch: true };
}

function documentAction(name, operationType, description) {
  return { name, operationType, description, batch: false };
}

const docxEditActions = [
  editAction('docx_set_anchored_text', 'replaceAnchoredText', 'Set text at current DOCX comment anchors.', z.object({ commentId: pathInput, text: z.string() }).strict()),
  editAction('docx_set_paragraph_text', 'replaceParagraphText', 'Set current body paragraph text.', z.object({ paragraphIndex: index, text: z.string() }).strict()),
  editAction('docx_set_paragraph_run_text', 'replaceParagraphRunText', 'Set current body paragraph run text.', z.object({ paragraphIndex: index, runIndex: index, text: z.string() }).strict()),
  editAction('docx_replace_body_text', 'replaceBodyText', 'Replace uniquely matched current body text.', z.object({ findText: pathInput, text: z.string() }).strict()),
  editAction('docx_delete_body_paragraph', 'deleteBodyParagraph', 'Delete uniquely matched current body paragraphs.', z.object({ findText: pathInput, ...optionalTextMatch }).strict()),
  editAction('docx_delete_body_drawing_before_paragraph', 'deleteBodyDrawingBeforeParagraph', 'Delete the drawing immediately before a uniquely matched current paragraph.', z.object({ findText: pathInput, ...optionalTextMatch }).strict()),
  editAction('docx_delete_body_range', 'deleteBodyRange', 'Delete uniquely bounded current body ranges.', z.object({ findText: pathInput, endFindText: z.string().optional(), matchMode: z.string().optional(), endMatchMode: z.string().optional(), paragraphStyle: z.string().optional(), endParagraphStyle: z.string().optional(), deleteToBodyEnd: z.boolean().optional(), removePrecedingPageBreak: z.boolean().optional() }).strict()),
  editAction('docx_start_section', 'startSectionBeforeParagraph', 'Start a section before a uniquely matched current paragraph.', z.object({ findText: pathInput, orientation: z.enum(['portrait', 'landscape']) }).strict()),
  editAction('docx_set_header_paragraph_text', 'replaceHeaderParagraphText', 'Set current header paragraph text.', z.object({ headerIndex: index, paragraphIndex: index, text: z.string() }).strict()),
  editAction('docx_set_header_run_text', 'replaceHeaderParagraphRunText', 'Set current header run text.', z.object({ headerIndex: index, paragraphIndex: index, runIndex: index, text: z.string() }).strict()),
  editAction('docx_replace_header_text', 'replaceHeaderText', 'Replace uniquely matched current header text.', z.object({ findText: pathInput, text: z.string() }).strict()),
  editAction('docx_set_footer_paragraph_text', 'replaceFooterParagraphText', 'Set current footer paragraph text.', z.object({ footerIndex: index, paragraphIndex: index, text: z.string() }).strict()),
  editAction('docx_set_footer_run_text', 'replaceFooterParagraphRunText', 'Set current footer run text.', z.object({ footerIndex: index, paragraphIndex: index, runIndex: index, text: z.string() }).strict()),
  editAction('docx_set_table_cell_text', 'replaceTableCellText', 'Set current body table cell text.', z.object({ tableIndex: index, rowIndex: index, cellIndex: index, text: z.string(), alignment: z.string().optional() }).strict()),
  editAction('docx_set_table_cell_run_text', 'replaceTableCellRunText', 'Set current body table cell run text.', z.object({ tableIndex: index, rowIndex: index, cellIndex: index, paragraphIndex: index, runIndex: index, text: z.string() }).strict()),
  editAction('docx_set_table_cell_choice_state', 'setTableCellChoiceState', 'Set the selected state represented by current table-cell content.', z.object({ tableIndex: index, rowIndex: index, cellIndex: index, text: z.string() }).strict()),
  editAction('docx_set_header_table_cell_text', 'replaceHeaderTableCellText', 'Set current header table cell text.', z.object({ headerIndex: index, tableIndex: index, rowIndex: index, cellIndex: index, text: z.string() }).strict()),
  editAction('docx_set_header_table_cell_run_text', 'replaceHeaderTableCellRunText', 'Set current header table cell run text.', z.object({ headerIndex: index, tableIndex: index, rowIndex: index, cellIndex: index, paragraphIndex: index, runIndex: index, text: z.string() }).strict()),
  editAction('docx_set_footer_table_cell_text', 'replaceFooterTableCellText', 'Set current footer table cell text.', z.object({ footerIndex: index, tableIndex: index, rowIndex: index, cellIndex: index, text: z.string() }).strict()),
  editAction('docx_set_footer_table_cell_run_text', 'replaceFooterTableCellRunText', 'Set current footer table cell run text.', z.object({ footerIndex: index, tableIndex: index, rowIndex: index, cellIndex: index, paragraphIndex: index, runIndex: index, text: z.string() }).strict()),
  editAction('docx_set_table_cell_rich_text', 'replaceTableCellRichText', 'Set current body table cell rich text.', z.object({ tableIndex: index, rowIndex: index, cellIndex: index, richText: z.array(richTextSegment) }).strict()),
  editAction('docx_insert_table_rows', 'insertTableRows', 'Insert rows into a current body table.', z.object({ tableIndex: index, rowIndex: index, templateRowIndex: index.optional(), rows: z.array(z.array(tableCellInput)) }).strict()),
  editAction('docx_delete_table_rows', 'deleteTableRows', 'Delete current body table row ranges.', z.object({ tableIndex: index, startRowIndex: index, endRowIndex: index }).strict()),
  editAction('docx_replace_table_rows', 'replaceTableRows', 'Replace current body table row ranges.', z.object({ tableIndex: index, startRowIndex: index, endRowIndex: index, templateRowIndex: index.optional(), rows: z.array(z.array(tableCellInput)) }).strict()),
  editAction('docx_insert_table_columns', 'insertTableColumns', 'Insert columns into a current body table.', z.object({ tableIndex: index, columnIndex: index, columnCount: positiveIndex.optional(), templateColumnIndex: index.optional() }).strict()),
  editAction('docx_set_table_width', 'setTableWidth', 'Set current body table widths.', z.object({ tableIndex: index, width: pathInput, widthType: z.enum(['pct', 'dxa', 'auto', 'nil']) }).strict()),
  editAction('docx_set_table_cell_alignment', 'setTableCellAlignment', 'Set current body table cell alignment.', z.object({ tableIndex: index, rowIndex: index, cellIndex: index, alignment: pathInput }).strict()),
  editAction('docx_set_table_cell_no_wrap', 'setTableCellNoWrap', 'Set current body table cell no-wrap state.', z.object({ tableIndex: index, rowIndex: index, cellIndex: index, noWrap: z.boolean() }).strict()),
  editAction('docx_set_table_cell_font_size', 'setTableCellFontSize', 'Set current body table cell font size.', z.object({ tableIndex: index, rowIndex: index, cellIndex: index, fontSize: pathInput }).strict()),
  editAction('docx_apply_font_policy', 'applyDocumentFontPolicy', 'Apply an explicit font policy to current document text.', z.object({ fontPolicy: z.object({ schema: pathInput, body: z.record(z.string(), z.string()), table: z.record(z.string(), z.string()) }).strict() }).strict()),
  editAction('docx_set_table_row_height', 'setTableRowHeight', 'Set current body table row height.', z.object({ tableIndex: index, rowIndex: index, height: pathInput, heightRule: z.string().optional() }).strict()),
  editAction('docx_set_table_row_cant_split', 'setTableRowCantSplit', 'Set current body table row split behavior.', z.object({ tableIndex: index, rowIndex: index, cantSplit: z.boolean() }).strict()),
  editAction('docx_set_table_row_keep_next', 'setTableRowKeepNext', 'Set keep-next behavior for current body table rows.', z.object({ tableIndex: index, rowIndex: index, keepNext: z.boolean() }).strict()),
  editAction('docx_set_body_paragraph_keep_next', 'setBodyParagraphKeepNext', 'Set keep-next behavior for current body paragraphs.', z.object({ paragraphIndex: index, keepNext: z.boolean() }).strict()),
  editAction('docx_set_body_paragraph_keep_lines', 'setBodyParagraphKeepLines', 'Set keep-lines behavior for current body paragraphs.', z.object({ paragraphIndex: index, keepLines: z.boolean() }).strict()),
  editAction('docx_set_header_paragraph_font_size', 'setHeaderParagraphFontSize', 'Set current header paragraph font size.', z.object({ headerIndex: index, paragraphIndex: index, fontSize: pathInput }).strict()),
  documentAction('docx_collapse_trailing_empty_section', 'collapseTrailingEmptySection', 'Collapse a current trailing empty section.'),
  documentAction('docx_collapse_trailing_empty_paragraphs', 'collapseTrailingEmptyBodyParagraphs', 'Collapse current trailing empty body paragraphs.'),
  editAction('docx_merge_table_cells', 'mergeTableCells', 'Merge current body table cells.', z.object({ tableIndex: index, rowIndex: index.optional(), startCellIndex: index.optional(), endCellIndex: index.optional(), startRowIndex: index.optional(), endRowIndex: index.optional(), cellIndex: index.optional(), gridColumn: index.optional() }).strict()),
  editAction('docx_unmerge_table_row_cells', 'unmergeTableRowHorizontalCells', 'Unmerge current horizontal table cells.', z.object({ tableIndex: index, rowIndex: index, cellIndex: index }).strict()),
  editAction('docx_unmerge_table_column_cells', 'unmergeTableColumnVerticalCells', 'Unmerge current vertical table cell ranges.', z.object({ tableIndex: index, cellIndex: index, startRowIndex: index, endRowIndex: index }).strict()),
  editAction('docx_delete_comments', 'deleteComments', 'Delete explicit current DOCX comments.', z.object({ commentIds: z.array(pathInput).min(1) }).strict()),
  documentAction('docx_mark_fields_dirty', 'markFieldsDirty', 'Mark current DOCX fields for native refresh.'),
  documentAction('docx_sanitize_fields', 'sanitizeFields', 'Remove update prompts and dirty markers from current DOCX fields.'),
  documentAction('docx_freeze_fields', 'freezeFields', 'Convert current visible DOCX field results to ordinary content.'),
];

const scalar = z.union([z.string(), z.number(), z.boolean(), z.null()]);
const xlsxEditActions = [
  editAction('xlsx_set_cell_value', 'setCellValue', 'Set current workbook cell values.', z.object({ sheet: pathInput, cell: pathInput, value: scalar, valueType: z.string().optional(), bold: z.boolean().optional(), shrinkToFit: z.boolean().optional(), wrapText: z.boolean().optional() }).strict()),
  editAction('xlsx_set_cell_number_format', 'setCellNumberFormat', 'Set current workbook cell number formats.', z.object({ sheet: pathInput, cell: pathInput, numberFormat: pathInput }).strict()),
  editAction('xlsx_set_rich_text_cell_value', 'setRichTextCellValue', 'Set current workbook rich-text cell values.', z.object({ sheet: pathInput, cell: pathInput, value: z.string(), bold: z.boolean() }).strict()),
  editAction('xlsx_set_range_values', 'setRangeValues', 'Set rectangular values in a current workbook.', z.object({ sheet: pathInput, startCell: pathInput, values: z.array(z.array(scalar)), valueType: z.string().optional() }).strict()),
  editAction('xlsx_insert_rows', 'insertRows', 'Insert rows into a current worksheet.', z.object({ sheet: pathInput, startRow: positiveIndex, count: positiveIndex, preserveHorizontalMergedRanges: z.boolean().optional(), expandAdjacentVerticalMergedRanges: z.boolean().optional() }).strict()),
  editAction('xlsx_copy_row', 'copyRow', 'Copy current worksheet rows.', z.object({ sheet: pathInput, sourceRow: positiveIndex, targetRow: positiveIndex, translateFormulas: z.boolean().optional() }).strict()),
  editAction('xlsx_expand_section_rows', 'expandSectionRows', 'Expand current worksheet row sections from visible anchors.', z.object({ sheet: pathInput, anchorText: pathInput, exampleRows: positiveIndex, targetRows: positiveIndex, preserveStyle: z.boolean().optional(), preserveFormulas: z.boolean().optional(), preserveMergedRanges: z.boolean().optional() }).strict()),
  editAction('xlsx_set_print_area', 'setPrintArea', 'Set current worksheet print areas.', z.object({ sheet: pathInput, range: pathInput }).strict()),
  editAction('xlsx_set_page_setup', 'setPageSetup', 'Set current worksheet page properties.', z.object({ sheet: pathInput, fitToPagesWide: positiveIndex.optional(), fitToPagesTall: positiveIndex.optional(), orientation: z.enum(['portrait', 'landscape']).optional(), paperSize: z.enum(['letter', 'legal', 'a3', 'a4']).optional(), repeatRowsStart: positiveIndex.optional(), repeatRowsEnd: positiveIndex.optional(), repeatColsStart: positiveIndex.optional(), repeatColsEnd: positiveIndex.optional() }).strict()),
  editAction('xlsx_set_row_page_breaks', 'setRowPageBreaks', 'Set current worksheet row page breaks.', z.object({ sheet: pathInput, breakBeforeRows: z.array(positiveIndex) }).strict()),
  editAction('xlsx_set_column_width', 'setColumnWidth', 'Set current worksheet column widths.', z.object({ sheet: pathInput, column: pathInput, width: z.number().positive().max(255) }).strict()),
];

function editToolDefinitions(actions) {
  return actions.map(action => ({
    name: action.name,
    description: action.batch ? `${action.description} One call batches only this action kind.` : action.description,
    inputSchema: z.object({
      input: pathInput,
      output: pathInput.describe('New document output path. Existing files are never overwritten.'),
      receiptOutput: pathInput.describe('New JSON receipt path. Existing files are never overwritten.'),
      ...(action.batch ? { changes: z.array(action.changeSchema).min(1) } : {}),
    }).strict(),
    outputSchema: fixedEditOutput(action.name),
    annotations: action.batch ? undefined : { idempotentHint: true },
    handler: args => fixedEdit(action, args),
  }));
}

function fixedEditOutput(tool) {
  return z.object({
    tool: z.literal(tool), runtime: runtimeIdentity, receipt: artifact, output: artifact.nullable(),
    summary: z.object({ pass: z.boolean(), operationCount: z.number().int().nonnegative(), appliedCount: z.number().int().nonnegative() }).strict(),
  }).strict();
}

function artifactOutput(tool) {
  return z.object({ tool: z.literal(tool), runtime: runtimeIdentity, artifact }).strict();
}

const tools = [
  {
    name: 'docx_inspect',
    description: 'Inspect a DOCX document and write one unified JSON observation containing placeholders, comments, anchors, tables, fields, flow, fonts, and formatting metrics.',
    inputSchema: artifactInput,
    outputSchema: artifactOutput('docx_inspect'),
    handler: docxInspect,
  },
  {
    name: 'docx_inspect_tables',
    description: 'Inspect current DOCX tables, cells, merges, paragraphs, runs, and formatting.',
    inputSchema: artifactInput,
    outputSchema: artifactOutput('docx_inspect_tables'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxInspectTables,
  },
  {
    name: 'docx_compare',
    description: 'Compare two DOCX files and report package, metric, and style differences.',
    inputSchema: z.object({ baseline: pathInput, updated: pathInput }).strict(),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxCompare,
  },
  {
    name: 'docx_export_json',
    description: 'Export DOCX body content to a new JSON artifact without returning the full document through MCP.',
    inputSchema: artifactInput,
    outputSchema: artifactOutput('docx_export_json'),
    handler: docxExportJson,
  },
  {
    name: 'docx_validate',
    description: 'Validate a current DOCX package against the published OpenXML contract.',
    inputSchema: inputOnly,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxValidate,
  },
  {
    name: 'docx_validate_font_policy',
    description: 'Validate current DOCX text against an explicit font policy.',
    inputSchema: z.object({ input: pathInput, policy: z.object({ schema: pathInput, body: z.record(z.string(), z.string()), table: z.record(z.string(), z.string()) }).strict() }).strict(),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxValidateFontPolicy,
  },
  {
    name: 'docx_strip_direct_formatting',
    description: 'Remove direct paragraph and run formatting while preserving styles.',
    inputSchema: z.object({ input: pathInput, output: pathInput }).strict(),
    handler: docxStripDirectFormatting,
  },
  {
    name: 'docx_replace_style_ids',
    description: 'Replace current DOCX style IDs from an explicit style map.',
    inputSchema: z.object({ input: pathInput, output: pathInput, styleMap: z.record(z.string(), z.string()) }).strict(),
    handler: docxReplaceStyleIds,
  },
  {
    name: 'docx_fill_template',
    description: 'Fill current DOCX placeholders from an explicit data object.',
    inputSchema: z.object({ template: pathInput, output: pathInput, data: z.record(z.string(), z.unknown()) }).strict(),
    handler: docxFillTemplate,
  },
  ...editToolDefinitions(docxEditActions),
  {
    name: 'office_render_pdf',
    description: 'Render a current Office document to PDF with its required native WPS backend and write the complete provider receipt as evidence. The input extension selects Writer, Spreadsheets, or Presentation; fallback rendering is rejected.',
    inputSchema: z.object({
      input: pathInput.describe('Path to the current Office document.'),
      output: pathInput.describe('New PDF output path. Existing files are never overwritten.'),
      receiptOutput: pathInput.describe('New JSON receipt path. Existing files are never overwritten.'),
    }).strict(),
    outputSchema: nativeRenderOutput,
    handler: officeRenderPdf,
  },
  {
    name: 'xlsx_convert_legacy',
    description: 'Convert a current legacy XLS workbook to XLSX using the published native ET backend.',
    inputSchema: z.object({ input: pathInput, output: pathInput, receiptOutput: pathInput }).strict(),
    handler: xlsxConvertLegacy,
  },
  {
    name: 'xlsx_inspect',
    description: 'Inspect a current XLSX workbook or legacy XLS workbook and write one JSON observation containing workbook structure, exported values, formulas, styles, merged ranges, and any published legacy-format conversion evidence.',
    inputSchema: artifactInput,
    outputSchema: artifactOutput('xlsx_inspect'),
    handler: xlsxInspect,
  },
  {
    name: 'xlsx_export_json',
    description: 'Export workbook sheet data from XLSX as structured JSON.',
    inputSchema: z.object({
      input: pathInput,
      output: pathInput.describe('New JSON artifact path. Existing files are never overwritten.'),
      resolveMergedCells: z.boolean().optional().describe('Resolve merged cells to project values.'),
    }).strict(),
    outputSchema: artifactOutput('xlsx_export_json'),
    handler: xlsxExportJson,
  },
  {
    name: 'xlsx_fill_template',
    description: 'Fill current XLSX placeholders from an explicit data object.',
    inputSchema: z.object({ template: pathInput, output: pathInput, data: z.record(z.string(), z.unknown()) }).strict(),
    handler: xlsxFillTemplate,
  },
  ...editToolDefinitions(xlsxEditActions),
  {
    name: 'xlsx_validate',
    description: 'Validate an XLSX workbook package and return Open XML validation evidence.',
    inputSchema: inputOnly,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: xlsxValidate,
  },
  {
    name: 'pptx_inspect',
    description: 'Inspect a PPTX file and write one detailed JSON observation containing slides, masters, layouts, shapes, transforms, paragraphs, runs, and placeholders.',
    inputSchema: artifactInput,
    outputSchema: artifactOutput('pptx_inspect'),
    handler: pptxInspect,
  },
  {
    name: 'pptx_export_json',
    description: 'Export PPTX slide text, notes, and placeholder hints to a new JSON artifact without returning the full presentation through MCP.',
    inputSchema: artifactInput,
    outputSchema: artifactOutput('pptx_export_json'),
    handler: pptxExportJson,
  },
  {
    name: 'pptx_fill_template',
    description: 'Fill current PPTX placeholders from an explicit data object.',
    inputSchema: z.object({ template: pathInput, output: pathInput, data: z.record(z.string(), z.unknown()) }).strict(),
    handler: pptxFillTemplate,
  },
  {
    name: 'pptx_apply_template',
    description: 'Apply one deterministic PPTX template-application plan to a current presentation. This tool executes the published plan; it does not select a template or derive business content, slide mappings, geometry, or formatting decisions.',
    inputSchema: z.object({
      input: pathInput.describe('Path to the current source PPTX.'),
      template: pathInput.describe('Path to the selected current template PPTX.'),
      targetMasterPath: pathInput,
      slides: z.array(z.object({ slideNumber: positiveIndex, targetLayoutPath: pathInput, contentBounds: z.object({ x: index, y: index, cx: positiveIndex, cy: positiveIndex }).strict().optional(), contentShapeIds: z.array(positiveIndex).min(1).optional(), sourceLayoutShapeIdsToPreserve: z.array(positiveIndex).optional() }).strict()).min(1),
      output: pathInput.describe('New PPTX output path. Existing files are never overwritten.'),
      receiptOutput: pathInput.describe('New JSON receipt path. Existing files are never overwritten.'),
    }).strict(),
    outputSchema: pptxTemplateApplyOutput,
    handler: pptxApplyTemplate,
  },
  {
    name: 'pptx_apply_format',
    description: 'Apply one deterministic PPTX formatting plan to a current presentation. This tool executes published formatting operations; it does not derive values, coordinates, or business decisions.',
    inputSchema: z.object({
      input: pathInput.describe('Path to the current PPTX.'),
      changes: z.array(z.object({ slideNumber: positiveIndex, shapeId: positiveIndex, runIndex: index, fontFamily: z.string().optional(), fontSize: z.number().positive().optional(), color: z.string().optional(), bold: z.boolean().optional(), paragraphAlignment: z.string().optional() }).strict()).min(1),
      output: pathInput.describe('New PPTX output path. Existing files are never overwritten.'),
      receiptOutput: pathInput.describe('New JSON receipt path. Existing files are never overwritten.'),
    }).strict(),
    outputSchema: pptxFormatApplyOutput,
    handler: pptxApplyFormat,
  },
  {
    name: 'pptx_validate',
    description: 'Validate a current PPTX package against the published OpenXML contract.',
    inputSchema: inputOnly,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: pptxValidate,
  },
];

function buildServer() {
  const server = new McpServer(
    { name: 'tiwater-office', version: packageMetadata.version },
    {
      instructions: 'Use these tools only for generic Office observation, conversion, editing, validation, and native rendering. Derive business meaning from the active scenario knowledge and current documents; the provider owns no scenario workflow.',
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
      async args => createToolResult(await tool.handler(args)),
    );
  }
  return server;
}

async function docxInspect(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(docxCandidates, ['inspect', input, '--json']);
  return {
    tool: 'docx_inspect',
    runtime: commandRuntime(result),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function docxInspectTables(args) {
  const result = await runJsonCandidateChain(docxCandidates, ['inspect-tables', requireString(args.input, 'input'), '--json']);
  return { tool: 'docx_inspect_tables', runtime: commandRuntime(result), artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json) };
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

async function docxStripDirectFormatting(args) {
  return copyTransform('docx_strip_direct_formatting', docxCandidates, ['strip-direct-formatting'], args);
}

async function docxReplaceStyleIds(args) {
  return withTempJsonFile(args.styleMap, styleMapPath => copyTransform('docx_replace_style_ids', docxCandidates, ['replace-style-ids'], args, [styleMapPath]));
}

async function docxFillTemplate(args) {
  return withTempJsonFile(args.data, dataPath => templateFill('docx_fill_template', docxCandidates, args, dataPath));
}

async function xlsxFillTemplate(args) {
  return withTempJsonFile(args.data, dataPath => templateFill('xlsx_fill_template', xlsxCandidates, args, dataPath));
}

async function pptxFillTemplate(args) {
  return withTempJsonFile(args.data, dataPath => templateFill('pptx_fill_template', pptxCandidates, args, dataPath));
}

async function templateFill(tool, candidates, args, dataPath) {
  const template = path.resolve(requireString(args.template, 'template'));
  const output = path.resolve(requireString(args.output, 'output'));
  await requireNewFile(output, 'output');
  const result = await runJsonCandidateChain(candidates, ['fill-template', template, dataPath, output]);
  return { tool, runtime: commandRuntime(result), output: await fileArtifact(output), result: result.json };
}

async function copyTransform(tool, candidates, command, args, suffix = []) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  await requireNewFile(output, 'output');
  const result = await runCandidateChain(candidates, [...command, input, output, ...suffix]);
  return { tool, runtime: commandRuntime(result), output: await fileArtifact(output) };
}

async function fixedEdit(action, args) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  const receiptOutput = path.resolve(requireString(args.receiptOutput, 'receiptOutput'));
  await requireNewFile(output, 'output');
  await requireNewFile(receiptOutput, 'receiptOutput');
  const inputArtifact = await fileArtifact(input);
  const operations = action.batch
    ? args.changes.map(change => ({ ...change, type: action.operationType }))
    : [{ type: action.operationType }];
  const candidates = action.name.startsWith('docx_') ? docxCandidates : xlsxCandidates;
  return withTempJsonFile({ operations }, async operationsPath => {
    try {
      const result = await runJsonCandidateChain(candidates, ['edit', input, operationsPath, output], { allowedExitCodes: [0, 1] });
      const appliedOperations = Array.isArray(result.json?.appliedOperations) ? result.json.appliedOperations : [];
      const pass = appliedOperations.length === operations.length && appliedOperations.every(operation => operation.applied === true);
      const outputArtifact = pass ? await fileArtifact(output) : null;
      if (!pass) await rm(output, { force: true });
      const receipt = {
        schema: 'tiwater.office.fixed-edit-receipt/v1',
        tool: action.name,
        operationType: action.operationType,
        pass,
        input: inputArtifact,
        output: outputArtifact,
        operationCount: operations.length,
        appliedOperations,
      };
      return {
        tool: action.name,
        runtime: commandRuntime(result),
        receipt: await writeJsonArtifact(receiptOutput, receipt),
        output: outputArtifact,
        summary: { pass, operationCount: operations.length, appliedCount: appliedOperations.filter(operation => operation.applied).length },
      };
    } catch (error) {
      await rm(output, { force: true });
      throw error;
    }
  });
}

async function docxCompare(args) {
  const baseline = requireString(args.baseline, 'baseline');
  const updated = requireString(args.updated, 'updated');
  const result = await runJsonCandidateChain(docxCandidates, ['compare', baseline, updated, '--json']);
  return { tool: 'docx_compare', runtime: commandRuntime(result), report: result.json };
}

async function docxExportJson(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(docxCandidates, ['export-json', input]);
  return {
    tool: 'docx_export_json',
    runtime: commandRuntime(result),
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
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(xlsxCandidates, ['inspect', input, '--json']);
  return {
    tool: 'xlsx_inspect',
    runtime: commandRuntime(result),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function xlsxExportJson(args) {
  const input = requireString(args.input, 'input');
  const cmdArgs = ['export-json', input];
  if (args.resolveMergedCells) {
    cmdArgs.push('--resolve-merged-cells');
  }
  const result = await runJsonCandidateChain(xlsxCandidates, cmdArgs);
  return {
    tool: 'xlsx_export_json',
    runtime: commandRuntime(result),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function xlsxValidate(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(xlsxCandidates, ['validate', input], { allowedExitCodes: [0, 1] });
  return { tool: 'xlsx_validate', runtime: commandRuntime(result), result: result.json };
}

async function pptxInspect(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(pptxCandidates, ['inspect', input, '--json']);
  return {
    tool: 'pptx_inspect',
    runtime: commandRuntime(result),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function pptxExportJson(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(pptxCandidates, ['export-json', input]);
  return {
    tool: 'pptx_export_json',
    runtime: commandRuntime(result),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
}

async function pptxApplyTemplate(args) {
  return withTempJsonFile({ targetMasterPath: args.targetMasterPath, slides: args.slides }, planPath => pptxApply('pptx_apply_template', args, true, planPath));
}

async function pptxApplyFormat(args) {
  return withTempJsonFile({ operations: args.changes }, planPath => pptxApply('pptx_apply_format', args, false, planPath));
}

async function pptxApply(tool, args, templateMode, plan) {
  const input = path.resolve(requireString(args.input, 'input'));
  const template = templateMode ? path.resolve(requireString(args.template, 'template')) : null;
  const output = path.resolve(requireString(args.output, 'output'));
  const receiptOutput = path.resolve(requireString(args.receiptOutput, 'receiptOutput'));
  for (const [label, candidate] of [['input', input], ['output', output], ...(template ? [['template', template]] : [])]) {
    if (path.extname(candidate).toLowerCase() !== '.pptx') {
      throw Object.assign(new Error(`${label} must use the .pptx extension`), { code: -32602 });
    }
  }
  await requireNewFile(output, 'output');
  await requireNewFile(receiptOutput, 'receiptOutput');
  const inputArtifact = await fileArtifact(input);
  const templateArtifact = template ? await fileArtifact(template) : null;
  const planArtifact = await fileArtifact(plan);
  await mkdir(path.dirname(output), { recursive: true });
  try {
    const command = templateMode ? 'apply-template' : 'apply-format-edits';
    const commandArgs = templateMode
      ? [command, input, template, plan, output]
      : [command, input, plan, output];
    const result = await runJsonCandidateChain(pptxCandidates, commandArgs, { allowedExitCodes: [0, 1] });
    await requireArtifactUnchanged(inputArtifact, 'PPTX apply input');
    if (templateArtifact) await requireArtifactUnchanged(templateArtifact, 'PPTX apply template');
    await requireArtifactUnchanged(planArtifact, 'PPTX apply plan');
    const providerResult = (templateMode ? pptxTemplateApplyResult : pptxFormatApplyResult).parse(result.json);
    if (path.resolve(providerResult.input) !== input
        || path.resolve(providerResult.output) !== output
        || (templateMode && path.resolve(providerResult.template) !== template)) {
      throw new Error('PPTX apply receipt is not bound to the current inputs and output');
    }
    const pass = providerResult.issues.length === 0;
    const outputArtifact = pass ? await fileArtifact(output) : null;
    if (!pass) await rm(output, { force: true });
    const receipt = {
      schema: templateMode
        ? 'tiwater.office.pptx-template-apply-receipt/v1'
        : 'tiwater.office.pptx-format-apply-receipt/v1',
      pass,
      input: inputArtifact,
      ...(templateMode ? { template: templateArtifact } : {}),
      requestSha256: planArtifact.sha256,
      output: outputArtifact,
      providerResult,
    };
    return {
      tool,
      runtime: commandRuntime(result),
      receipt: await writeJsonArtifact(receiptOutput, receipt),
      output: outputArtifact,
      summary: templateMode
        ? { pass, changedSlideCount: providerResult.changedSlideCount, issueCount: providerResult.issues.length }
        : { pass, operationCount: providerResult.operationCount, changedCount: providerResult.changedCount, issueCount: providerResult.issues.length },
    };
  } catch (error) {
    await rm(output, { force: true });
    throw error;
  }
}

async function pptxValidate(args) {
  const result = await runJsonCandidateChain(pptxCandidates, ['validate', requireString(args.input, 'input')], { allowedExitCodes: [0, 1] });
  return { tool: 'pptx_validate', runtime: commandRuntime(result), result: result.json };
}

async function requireArtifactUnchanged(expected, label) {
  const current = await fileArtifact(expected.path);
  if (!isDeepStrictEqual(current, expected)) throw new Error(`${label} changed during provider execution`);
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
