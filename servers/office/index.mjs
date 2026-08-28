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
const pptxObjectIssue = z.object({
  slideNumber: z.number().int(),
  shapeId: z.number().int().nonnegative().max(0xffffffff),
  message: z.string().min(1),
}).strict();
const pptxTransform = z.object({
  x: z.number().int(), y: z.number().int(), cx: z.number().int().positive(), cy: z.number().int().positive(),
}).strict();
const pptxShapeGeometryResult = z.object({
  input: z.string().min(1), output: z.string().min(1),
  operationCount: z.number().int().nonnegative(), appliedCount: z.number().int().nonnegative(),
  changes: z.array(z.object({
    slideNumber: z.number().int().positive(), shapeId: z.number().int().positive().max(0xffffffff),
    before: pptxTransform, after: pptxTransform,
  }).strict()),
  issues: z.array(pptxObjectIssue),
}).strict();
const pptxPictureImageResult = z.object({
  input: z.string().min(1), output: z.string().min(1),
  operationCount: z.number().int().nonnegative(), appliedCount: z.number().int().nonnegative(),
  changes: z.array(z.object({
    slideNumber: z.number().int().positive(), shapeId: z.number().int().positive().max(0xffffffff), image: z.string().min(1),
    beforeSha256: z.string().regex(/^[0-9a-f]{64}$/), afterSha256: z.string().regex(/^[0-9a-f]{64}$/),
  }).strict()),
  issues: z.array(pptxObjectIssue),
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

const docxEditActions = [
  {"name":"docx_set_anchored_text","operationType":"replaceAnchoredText","description":"Set text at current DOCX comment anchors.","batch":true},
  {"name":"docx_set_paragraph_text","operationType":"replaceParagraphText","description":"Set current body paragraph text.","batch":true},
  {"name":"docx_set_paragraph_run_text","operationType":"replaceParagraphRunText","description":"Set current body paragraph run text.","batch":true},
  {"name":"docx_replace_body_text","operationType":"replaceBodyText","description":"Replace uniquely matched current body text.","batch":true},
  {"name":"docx_delete_body_paragraph","operationType":"deleteBodyParagraph","description":"Delete uniquely matched current body paragraphs.","batch":true},
  {"name":"docx_delete_body_drawing_before_paragraph","operationType":"deleteBodyDrawingBeforeParagraph","description":"Delete the drawing immediately before a uniquely matched current paragraph.","batch":true},
  {"name":"docx_insert_body_range","operationType":"insertBodyRange","description":"Insert a bounded direct-body range from a current source DOCX before a current target body boundary, preserving supported styles and relationships.","batch":true,"sourceFields":["source"]},
  {"name":"docx_replace_drawing_image","operationType":"replaceDrawingImage","description":"Replace the image relationship of a current body drawing while preserving its drawing geometry.","batch":true,"sourceFields":["image"]},
  {"name":"docx_insert_body_image","operationType":"insertBodyImage","description":"Insert an image as a new inline drawing before a current direct-body boundary.","batch":true,"sourceFields":["image"]},
  {"name":"docx_delete_body_range","operationType":"deleteBodyRange","description":"Delete uniquely bounded current body ranges.","batch":true},
  {"name":"docx_start_section","operationType":"startSectionBeforeParagraph","description":"Start a section before a uniquely matched current paragraph.","batch":true},
  {"name":"docx_set_header_paragraph_text","operationType":"replaceHeaderParagraphText","description":"Set current header paragraph text.","batch":true},
  {"name":"docx_set_header_run_text","operationType":"replaceHeaderParagraphRunText","description":"Set current header run text.","batch":true},
  {"name":"docx_replace_header_text","operationType":"replaceHeaderText","description":"Replace uniquely matched current header text.","batch":true},
  {"name":"docx_set_footer_paragraph_text","operationType":"replaceFooterParagraphText","description":"Set current footer paragraph text.","batch":true},
  {"name":"docx_set_footer_run_text","operationType":"replaceFooterParagraphRunText","description":"Set current footer run text.","batch":true},
  {"name":"docx_set_table_cell_text","operationType":"replaceTableCellText","description":"Set current body table cell text.","batch":true},
  {"name":"docx_set_table_cell_run_text","operationType":"replaceTableCellRunText","description":"Set current body table cell run text.","batch":true},
  {"name":"docx_set_header_table_cell_text","operationType":"replaceHeaderTableCellText","description":"Set current header table cell text.","batch":true},
  {"name":"docx_set_header_table_cell_run_text","operationType":"replaceHeaderTableCellRunText","description":"Set current header table cell run text.","batch":true},
  {"name":"docx_set_footer_table_cell_text","operationType":"replaceFooterTableCellText","description":"Set current footer table cell text.","batch":true},
  {"name":"docx_set_footer_table_cell_run_text","operationType":"replaceFooterTableCellRunText","description":"Set current footer table cell run text.","batch":true},
  {"name":"docx_set_table_cell_rich_text","operationType":"replaceTableCellRichText","description":"Set current body table cell rich text.","batch":true},
  {"name":"docx_insert_table_rows","operationType":"insertTableRows","description":"Insert rows into a current body table.","batch":true},
  {"name":"docx_delete_table_rows","operationType":"deleteTableRows","description":"Delete current body table row ranges.","batch":true},
  {"name":"docx_replace_table_rows","operationType":"replaceTableRows","description":"Replace current body table row ranges.","batch":true},
  {"name":"docx_insert_table_columns","operationType":"insertTableColumns","description":"Insert columns into a current body table.","batch":true},
  {"name":"docx_set_table_width","operationType":"setTableWidth","description":"Set current body table widths.","batch":true},
  {"name":"docx_set_table_cell_alignment","operationType":"setTableCellAlignment","description":"Set current body table cell alignment.","batch":true},
  {"name":"docx_set_table_cell_no_wrap","operationType":"setTableCellNoWrap","description":"Set current body table cell no-wrap state.","batch":true},
  {"name":"docx_set_table_cell_font_size","operationType":"setTableCellFontSize","description":"Set current body table cell font size.","batch":true},
  {"name":"docx_apply_font_policy","operationType":"applyDocumentFontPolicy","description":"Apply an explicit font policy to current document text.","batch":true},
  {"name":"docx_set_table_row_height","operationType":"setTableRowHeight","description":"Set current body table row height.","batch":true},
  {"name":"docx_set_table_row_cant_split","operationType":"setTableRowCantSplit","description":"Set current body table row split behavior.","batch":true},
  {"name":"docx_set_table_row_repeat_as_header","operationType":"setTableRowRepeatAsHeader","description":"Set or unset repeat-as-header on uniquely addressed current body, header, or footer table rows.","batch":true},
  {"name":"docx_set_table_row_keep_next","operationType":"setTableRowKeepNext","description":"Set keep-next behavior for current body table rows.","batch":true},
  {"name":"docx_set_body_paragraph_keep_next","operationType":"setBodyParagraphKeepNext","description":"Set keep-next behavior for current body paragraphs.","batch":true},
  {"name":"docx_set_body_paragraph_keep_lines","operationType":"setBodyParagraphKeepLines","description":"Set keep-lines behavior for current body paragraphs.","batch":true},
  {"name":"docx_apply_toc_style_policy","operationType":"applyTocStylePolicy","description":"Apply current document table-of-contents paragraph style properties.","batch":true},
  {"name":"docx_set_header_paragraph_font_size","operationType":"setHeaderParagraphFontSize","description":"Set current header paragraph font size.","batch":true},
  {"name":"docx_collapse_trailing_empty_section","operationType":"collapseTrailingEmptySection","description":"Collapse a current trailing empty section.","batch":false},
  {"name":"docx_collapse_trailing_empty_paragraphs","operationType":"collapseTrailingEmptyBodyParagraphs","description":"Collapse current trailing empty body paragraphs.","batch":false},
  {"name":"docx_merge_table_cells","operationType":"mergeTableCells","description":"Merge current body table cells.","batch":true},
  {"name":"docx_unmerge_table_row_cells","operationType":"unmergeTableRowHorizontalCells","description":"Unmerge current horizontal table cells.","batch":true},
  {"name":"docx_unmerge_table_column_cells","operationType":"unmergeTableColumnVerticalCells","description":"Unmerge current vertical table cell ranges.","batch":true},
  {"name":"docx_delete_comments","operationType":"deleteComments","description":"Delete explicit current DOCX comments.","batch":true},
  {"name":"docx_mark_fields_dirty","operationType":"markFieldsDirty","description":"Mark current DOCX fields for native refresh.","batch":false},
  {"name":"docx_sanitize_fields","operationType":"sanitizeFields","description":"Remove update prompts and dirty markers from current DOCX fields.","batch":false},
  {"name":"docx_freeze_fields","operationType":"freezeFields","description":"Convert current visible DOCX field results to ordinary content.","batch":false},
];

const xlsxEditActions = [
  {"name":"xlsx_set_cell_value","operationType":"setCellValue","description":"Set current workbook cell values.","batch":true},
  {"name":"xlsx_set_cell_number_format","operationType":"setCellNumberFormat","description":"Set current workbook cell number formats.","batch":true},
  {"name":"xlsx_set_rich_text_cell_value","operationType":"setRichTextCellValue","description":"Set current workbook rich-text cell values.","batch":true},
  {"name":"xlsx_set_range_values","operationType":"setRangeValues","description":"Set rectangular values in a current workbook.","batch":true},
  {"name":"xlsx_insert_rows","operationType":"insertRows","description":"Insert rows into a current worksheet.","batch":true},
  {"name":"xlsx_delete_rows","operationType":"deleteRows","description":"Structurally delete rows from a current worksheet.","batch":true},
  {"name":"xlsx_copy_row","operationType":"copyRow","description":"Copy current worksheet rows.","batch":true},
  {"name":"xlsx_expand_section_rows","operationType":"expandSectionRows","description":"Expand current worksheet row sections from visible anchors.","batch":true},
  {"name":"xlsx_set_print_area","operationType":"setPrintArea","description":"Set current worksheet print areas.","batch":true},
  {"name":"xlsx_set_page_setup","operationType":"setPageSetup","description":"Set current worksheet page properties.","batch":true},
  {"name":"xlsx_set_row_page_breaks","operationType":"setRowPageBreaks","description":"Set current worksheet row page breaks.","batch":true},
  {"name":"xlsx_set_column_width","operationType":"setColumnWidth","description":"Set current worksheet column widths.","batch":true},
];

function editToolDefinitions(actions) {
  return actions.map(action => ({
    name: action.name,
    description: action.batch ? `${action.description} One call batches only this action kind.` : action.description,
    inputSchema: inputContract(action.name),
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
    inputSchema: inputContract('docx_inspect'),
    outputSchema: artifactOutput('docx_inspect'),
    handler: docxInspect,
  },
  {
    name: 'docx_inspect_tables',
    description: 'Inspect current DOCX tables, cells, merges, paragraphs, runs, and formatting.',
    inputSchema: inputContract('docx_inspect_tables'),
    outputSchema: artifactOutput('docx_inspect_tables'),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxInspectTables,
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
    description: 'Export DOCX body content to a new JSON artifact without returning the full document through MCP.',
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
  ...editToolDefinitions(docxEditActions),
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
  ...editToolDefinitions(xlsxEditActions),
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
    outputSchema: pptxTemplateApplyOutput,
    handler: pptxApplyTemplate,
  },
  {
    name: 'pptx_apply_format',
    description: 'Apply one deterministic PPTX formatting plan to a current presentation. This tool executes published formatting operations; it does not derive values, coordinates, or business decisions.',
    inputSchema: inputContract('pptx_apply_format'),
    outputSchema: pptxFormatApplyOutput,
    handler: pptxApplyFormat,
  },
  {
    name: 'pptx_set_shape_geometry',
    description: 'Set exact native EMU bounds for uniquely identified current-slide PPTX objects. One call batches only this fixed geometry action and does not infer repair coordinates.',
    inputSchema: inputContract('pptx_set_shape_geometry'),
    outputSchema: fixedEditOutput('pptx_set_shape_geometry'),
    handler: pptxSetShapeGeometry,
  },
  {
    name: 'pptx_replace_picture_image',
    description: 'Replace embedded PNG or JPEG media for uniquely identified current-slide PPTX pictures while preserving the picture object, geometry, crop, and unrelated media. One call batches only this fixed replacement action.',
    inputSchema: inputContract('pptx_replace_picture_image'),
    outputSchema: fixedEditOutput('pptx_replace_picture_image'),
    handler: pptxReplacePictureImage,
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
  const sourcePaths = [...new Set((action.sourceFields ?? []).flatMap(field =>
    (args.changes ?? []).map(change => path.resolve(requireString(change[field], field)))))];
  const sources = await Promise.all(sourcePaths.map(fileArtifact));
  const candidates = action.name.startsWith('docx_') ? docxCandidates : xlsxCandidates;
  return withTempJsonFile({ operations }, async operationsPath => {
    try {
      const result = await runJsonCandidateChain(candidates, ['edit', input, operationsPath, output], { allowedExitCodes: [0, 1] });
      const rawAppliedOperations = result.json?.appliedOperations ?? result.json?.AppliedOperations;
      const appliedOperations = Array.isArray(rawAppliedOperations)
        ? rawAppliedOperations.map(operation => ({
            type: operation.type ?? operation.Type,
            applied: operation.applied ?? operation.Applied,
            detail: operation.detail ?? operation.Detail,
          }))
        : [];
      const observedSources = await Promise.all(sourcePaths.map(fileArtifact));
      const sourceBindingStable = isDeepStrictEqual(sources, observedSources);
      const pass = sourceBindingStable && appliedOperations.length === operations.length && appliedOperations.every(operation => operation.applied === true);
      const outputArtifact = pass ? await fileArtifact(output) : null;
      if (!pass) await rm(output, { force: true });
      const receipt = {
        schema: 'tiwater.office.fixed-edit-receipt/v1',
        tool: action.name,
        operationType: action.operationType,
        pass,
        input: inputArtifact,
        ...(sources.length > 0 ? { sources } : {}),
        ...(sources.length > 0 ? { sourceBindingStable } : {}),
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

async function pptxSetShapeGeometry(args) {
  return pptxFixedObjectEdit('pptx_set_shape_geometry', 'set-shape-geometry', args, pptxShapeGeometryResult, args.changes);
}

async function pptxReplacePictureImage(args) {
  const changes = args.changes.map(change => ({ ...change, image: path.resolve(requireString(change.image, 'image')) }));
  return pptxFixedObjectEdit('pptx_replace_picture_image', 'replace-picture-image', args, pptxPictureImageResult, changes, changes.map(change => change.image));
}

async function pptxFixedObjectEdit(tool, command, args, resultSchema, changes, sourcePaths = []) {
  const input = path.resolve(requireString(args.input, 'input'));
  const output = path.resolve(requireString(args.output, 'output'));
  const receiptOutput = path.resolve(requireString(args.receiptOutput, 'receiptOutput'));
  if (path.extname(input).toLowerCase() !== '.pptx' || path.extname(output).toLowerCase() !== '.pptx')
    throw Object.assign(new Error('PPTX object edits require .pptx input and output paths'), { code: -32602 });
  await requireNewFile(output, 'output');
  await requireNewFile(receiptOutput, 'receiptOutput');
  const inputArtifact = await fileArtifact(input);
  const sourceArtifacts = await Promise.all([...new Set(sourcePaths)].map(fileArtifact));
  return withTempJsonFile({ changes }, async planPath => {
    const requestArtifact = await fileArtifact(planPath);
    try {
      const result = await runJsonCandidateChain(pptxCandidates, [command, input, planPath, output], { allowedExitCodes: [0, 1] });
      await requireArtifactUnchanged(inputArtifact, 'PPTX object edit input');
      await requireArtifactUnchanged(requestArtifact, 'PPTX object edit request');
      for (const source of sourceArtifacts) await requireArtifactUnchanged(source, 'PPTX replacement image');
      const providerResult = resultSchema.parse(result.json);
      if (path.resolve(providerResult.input) !== input || path.resolve(providerResult.output) !== output)
        throw new Error('PPTX object edit receipt is not bound to the current input and output');
      const sourceByPath = new Map(sourceArtifacts.map(source => [source.path, source]));
      const providerMatchesRequest = providerResult.changes.length === changes.length && providerResult.changes.every((change, position) => {
        const requested = changes[position];
        if (change.slideNumber !== requested.slideNumber || change.shapeId !== requested.shapeId) return false;
        if (tool === 'pptx_set_shape_geometry')
          return isDeepStrictEqual(change.after, { x: requested.x, y: requested.y, cx: requested.cx, cy: requested.cy });
        const requestedImage = path.resolve(requested.image);
        return path.resolve(change.image) === requestedImage && change.afterSha256 === sourceByPath.get(requestedImage)?.sha256;
      });
      const pass = providerResult.issues.length === 0
        && providerResult.operationCount === changes.length
        && providerResult.appliedCount === changes.length
        && providerMatchesRequest;
      const outputArtifact = pass ? await fileArtifact(output) : null;
      if (!pass) await rm(output, { force: true });
      const receipt = {
        schema: 'tiwater.office.pptx-fixed-object-edit-receipt/v1', tool, pass,
        input: inputArtifact, requestSha256: requestArtifact.sha256,
        ...(sourceArtifacts.length ? { sourceImages: sourceArtifacts } : {}),
        output: outputArtifact, providerResult,
      };
      return {
        tool, runtime: commandRuntime(result), receipt: await writeJsonArtifact(receiptOutput, receipt), output: outputArtifact,
        summary: { pass, operationCount: providerResult.operationCount, appliedCount: providerResult.appliedCount },
      };
    } catch (error) {
      await rm(output, { force: true });
      throw error;
    }
  });
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
