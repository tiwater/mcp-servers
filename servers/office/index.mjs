#!/usr/bin/env node
import path from 'node:path';
import { spawn } from 'node:child_process';
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
const pptxProject = resolveRepoPath('packages', 'pptx-cli', 'pptx.csproj');

const docxCandidates = [
  commandCandidate('tiwater-docx'),
  commandCandidate('dotnet', ['run', '--project', docxProject, '--']),
];

const xlsxCandidates = [
  commandCandidate('tiwater-xlsx'),
  commandCandidate('dotnet', ['run', '--project', xlsxProject, '--']),
];

const pptxCandidates = [
  commandCandidate('tiwater-pptx'),
  commandCandidate('dotnet', ['run', '--project', pptxProject, '--']),
];

const integer = { type: 'integer', minimum: 0 };
const string = { type: 'string' };
const boolean = { type: 'boolean' };
const richText = {
  type: 'array',
  items: {
    type: 'object',
    properties: { text: string, color: string, underline: boolean, bold: boolean, fontName: string },
    required: ['text'],
    additionalProperties: false,
  },
};
const tableRows = { type: 'array', items: { type: 'array', items: { type: 'object' } } };

const docxEditActions = [
  action('docx_set_anchored_text', 'replaceAnchoredText', 'Set text at a current DOCX comment anchor.', { commentId: string, text: string }, ['commentId', 'text']),
  action('docx_set_paragraph_text', 'replaceParagraphText', 'Set a body paragraph text by current paragraph index.', { paragraphIndex: integer, text: string }, ['paragraphIndex', 'text']),
  action('docx_set_paragraph_run_text', 'replaceParagraphRunText', 'Set a run text in a current body paragraph.', { paragraphIndex: integer, runIndex: integer, text: string }, ['paragraphIndex', 'runIndex', 'text']),
  action('docx_replace_body_text', 'replaceBodyText', 'Replace unique current body text.', { findText: string, text: string }, ['findText', 'text']),
  action('docx_delete_body_paragraph', 'deleteBodyParagraph', 'Delete a uniquely selected current body paragraph.', { findText: string, matchMode: string, paragraphStyle: string }, ['findText']),
  action('docx_delete_body_range', 'deleteBodyRange', 'Delete a uniquely bounded current body range.', { findText: string, endFindText: string, matchMode: string, endMatchMode: string, paragraphStyle: string, endParagraphStyle: string, deleteToBodyEnd: boolean, removePrecedingPageBreak: boolean }, ['findText']),
  action('docx_start_section', 'startSectionBeforeParagraph', 'Start a section before a uniquely selected current paragraph.', { findText: string, orientation: { type: 'string', enum: ['portrait', 'landscape'] } }, ['findText', 'orientation']),
  action('docx_set_header_paragraph_text', 'replaceHeaderParagraphText', 'Set text in a current header paragraph.', { headerIndex: integer, paragraphIndex: integer, text: string }, ['headerIndex', 'paragraphIndex', 'text']),
  action('docx_set_header_run_text', 'replaceHeaderParagraphRunText', 'Set text in a current header paragraph run.', { headerIndex: integer, paragraphIndex: integer, runIndex: integer, text: string }, ['headerIndex', 'paragraphIndex', 'runIndex', 'text']),
  action('docx_replace_header_text', 'replaceHeaderText', 'Replace unique current text inside headers.', { findText: string, text: string }, ['findText', 'text']),
  action('docx_set_footer_paragraph_text', 'replaceFooterParagraphText', 'Set text in a current footer paragraph.', { footerIndex: integer, paragraphIndex: integer, text: string }, ['footerIndex', 'paragraphIndex', 'text']),
  action('docx_set_footer_run_text', 'replaceFooterParagraphRunText', 'Set text in a current footer paragraph run.', { footerIndex: integer, paragraphIndex: integer, runIndex: integer, text: string }, ['footerIndex', 'paragraphIndex', 'runIndex', 'text']),
  action('docx_set_table_cell_text', 'replaceTableCellText', 'Set text in a current body table cell.', { tableIndex: integer, rowIndex: integer, cellIndex: integer, text: string, alignment: string }, ['tableIndex', 'rowIndex', 'cellIndex', 'text']),
  action('docx_set_table_cell_run_text', 'replaceTableCellRunText', 'Set a run text in a current body table cell.', { tableIndex: integer, rowIndex: integer, cellIndex: integer, paragraphIndex: integer, runIndex: integer, text: string }, ['tableIndex', 'rowIndex', 'cellIndex', 'paragraphIndex', 'runIndex', 'text']),
  action('docx_set_header_table_cell_text', 'replaceHeaderTableCellText', 'Set text in a current header table cell.', { headerIndex: integer, tableIndex: integer, rowIndex: integer, cellIndex: integer, text: string }, ['headerIndex', 'tableIndex', 'rowIndex', 'cellIndex', 'text']),
  action('docx_set_header_table_cell_run_text', 'replaceHeaderTableCellRunText', 'Set a run text in a current header table cell.', { headerIndex: integer, tableIndex: integer, rowIndex: integer, cellIndex: integer, paragraphIndex: integer, runIndex: integer, text: string }, ['headerIndex', 'tableIndex', 'rowIndex', 'cellIndex', 'paragraphIndex', 'runIndex', 'text']),
  action('docx_set_footer_table_cell_text', 'replaceFooterTableCellText', 'Set text in a current footer table cell.', { footerIndex: integer, tableIndex: integer, rowIndex: integer, cellIndex: integer, text: string }, ['footerIndex', 'tableIndex', 'rowIndex', 'cellIndex', 'text']),
  action('docx_set_footer_table_cell_run_text', 'replaceFooterTableCellRunText', 'Set a run text in a current footer table cell.', { footerIndex: integer, tableIndex: integer, rowIndex: integer, cellIndex: integer, paragraphIndex: integer, runIndex: integer, text: string }, ['footerIndex', 'tableIndex', 'rowIndex', 'cellIndex', 'paragraphIndex', 'runIndex', 'text']),
  action('docx_set_table_cell_rich_text', 'replaceTableCellRichText', 'Set rich text in a current body table cell.', { tableIndex: integer, rowIndex: integer, cellIndex: integer, richText }, ['tableIndex', 'rowIndex', 'cellIndex', 'richText']),
  action('docx_insert_table_rows', 'insertTableRows', 'Insert rows into a current body table.', { tableIndex: integer, rowIndex: integer, templateRowIndex: integer, rows: tableRows }, ['tableIndex', 'rowIndex', 'rows']),
  action('docx_delete_table_rows', 'deleteTableRows', 'Delete a current body table row range.', { tableIndex: integer, startRowIndex: integer, endRowIndex: integer }, ['tableIndex', 'startRowIndex', 'endRowIndex']),
  action('docx_replace_table_rows', 'replaceTableRows', 'Replace a current body table row range.', { tableIndex: integer, startRowIndex: integer, endRowIndex: integer, templateRowIndex: integer, rows: tableRows }, ['tableIndex', 'startRowIndex', 'endRowIndex', 'rows']),
  action('docx_insert_table_columns', 'insertTableColumns', 'Insert columns into a current body table.', { tableIndex: integer, columnIndex: integer, columnCount: integer, templateColumnIndex: integer }, ['tableIndex', 'columnIndex']),
  action('docx_set_table_width', 'setTableWidth', 'Set a current body table width.', { tableIndex: integer, width: string, widthType: { type: 'string', enum: ['pct', 'dxa', 'auto', 'nil'] } }, ['tableIndex', 'width', 'widthType']),
  action('docx_set_table_cell_alignment', 'setTableCellAlignment', 'Set alignment for a current body table cell.', { tableIndex: integer, rowIndex: integer, cellIndex: integer, alignment: string }, ['tableIndex', 'rowIndex', 'cellIndex', 'alignment']),
  action('docx_set_table_cell_no_wrap', 'setTableCellNoWrap', 'Set no-wrap for a current body table cell.', { tableIndex: integer, rowIndex: integer, cellIndex: integer, noWrap: boolean }, ['tableIndex', 'rowIndex', 'cellIndex', 'noWrap']),
  action('docx_set_table_cell_font_size', 'setTableCellFontSize', 'Set font size for a current body table cell.', { tableIndex: integer, rowIndex: integer, cellIndex: integer, fontSize: string }, ['tableIndex', 'rowIndex', 'cellIndex', 'fontSize']),
  action('docx_set_table_row_height', 'setTableRowHeight', 'Set a current body table row height.', { tableIndex: integer, rowIndex: integer, height: string, heightRule: string }, ['tableIndex', 'rowIndex', 'height']),
  action('docx_set_table_row_cant_split', 'setTableRowCantSplit', 'Set whether a current body table row may split across pages.', { tableIndex: integer, rowIndex: integer, cantSplit: boolean }, ['tableIndex', 'rowIndex', 'cantSplit']),
  action('docx_set_table_row_keep_next', 'setTableRowKeepNext', 'Set keep-next behavior for a current body table row.', { tableIndex: integer, rowIndex: integer, keepNext: boolean }, ['tableIndex', 'rowIndex', 'keepNext']),
  action('docx_apply_font_policy', 'applyDocumentFontPolicy', 'Apply one explicit font policy to current body and table text.', { fontPolicy: { type: 'object' } }, ['fontPolicy']),
  onceAction('docx_collapse_trailing_empty_section', 'collapseTrailingEmptySection', 'Collapse a current trailing empty section.'),
  action('docx_merge_table_cells', 'mergeTableCells', 'Merge a current body table cell range.', { tableIndex: integer, rowIndex: integer, startCellIndex: integer, endCellIndex: integer, startRowIndex: integer, endRowIndex: integer, cellIndex: integer, gridColumn: integer }, ['tableIndex']),
  action('docx_unmerge_table_row_cells', 'unmergeTableRowHorizontalCells', 'Unmerge one current horizontal table cell.', { tableIndex: integer, rowIndex: integer, cellIndex: integer }, ['tableIndex', 'rowIndex', 'cellIndex']),
  action('docx_unmerge_table_column_cells', 'unmergeTableColumnVerticalCells', 'Unmerge one current vertical table cell range.', { tableIndex: integer, cellIndex: integer, startRowIndex: integer, endRowIndex: integer }, ['tableIndex', 'cellIndex', 'startRowIndex', 'endRowIndex']),
  action('docx_delete_comments', 'deleteComments', 'Delete an explicit set of current DOCX comments.', { commentIds: { type: 'array', minItems: 1, items: string } }, ['commentIds']),
  onceAction('docx_mark_fields_dirty', 'markFieldsDirty', 'Mark current DOCX fields for native refresh.'),
  onceAction('docx_sanitize_fields', 'sanitizeFields', 'Remove update prompts and dirty markers from current DOCX fields.'),
  onceAction('docx_freeze_fields', 'freezeFields', 'Convert current visible DOCX field results to ordinary content.'),
];

const xlsxEditActions = [
  action('xlsx_set_cell_value', 'setCellValue', 'Set values in current workbook cells.', { sheet: string, cell: string, value: {}, valueType: string, bold: boolean, shrinkToFit: boolean, wrapText: boolean }, ['sheet', 'cell', 'value']),
  action('xlsx_set_range_values', 'setRangeValues', 'Set rectangular values in a current workbook.', { sheet: string, startCell: string, values: { type: 'array', items: { type: 'array', items: {} } }, valueType: string }, ['sheet', 'startCell', 'values']),
  action('xlsx_insert_rows', 'insertRows', 'Insert rows into a current worksheet.', { sheet: string, startRow: { type: 'integer', minimum: 1 }, count: { type: 'integer', minimum: 1 } }, ['sheet', 'startRow', 'count']),
  action('xlsx_copy_row', 'copyRow', 'Copy one current worksheet row.', { sheet: string, sourceRow: { type: 'integer', minimum: 1 }, targetRow: { type: 'integer', minimum: 1 }, translateFormulas: boolean }, ['sheet', 'sourceRow', 'targetRow']),
  action('xlsx_expand_section_rows', 'expandSectionRows', 'Expand a current worksheet row section from its visible anchor.', { sheet: string, anchorText: string, exampleRows: { type: 'integer', minimum: 1 }, targetRows: { type: 'integer', minimum: 1 }, preserveStyle: boolean, preserveFormulas: boolean, preserveMergedRanges: boolean }, ['sheet', 'anchorText', 'exampleRows', 'targetRows']),
  action('xlsx_set_print_area', 'setPrintArea', 'Set a current worksheet print area.', { sheet: string, range: string }, ['sheet', 'range']),
  action('xlsx_set_page_setup', 'setPageSetup', 'Set current worksheet page properties.', { sheet: string, fitToPagesWide: { type: 'integer', minimum: 1 }, fitToPagesTall: { type: 'integer', minimum: 1 }, orientation: { type: 'string', enum: ['portrait', 'landscape'] }, paperSize: { type: 'string', enum: ['letter', 'legal', 'a3', 'a4'] } }, ['sheet']),
  action('xlsx_set_column_width', 'setColumnWidth', 'Set a current worksheet column width.', { sheet: string, column: string, width: { type: 'number', exclusiveMinimum: 0, maximum: 255 } }, ['sheet', 'column', 'width']),
];

function action(name, operationType, description, properties, required) {
  return { name, operationType, description, properties, required, batch: true };
}

function onceAction(name, operationType, description) {
  return { name, operationType, description, batch: false };
}

function editToolDefinitions(actions) {
  return actions.map(({ name, description, properties, required, batch }) => ({
    name,
    description: batch ? `${description} A call may batch only this one action kind.` : description,
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
        ...(batch ? { changes: { type: 'array', minItems: 1, items: { type: 'object', properties, required, additionalProperties: false } } } : {}),
      },
      required: batch ? ['input', 'output', 'changes'] : ['input', 'output'],
      additionalProperties: false,
    },
  }));
}

const docxEditActionByName = new Map(docxEditActions.map(value => [value.name, value]));
const xlsxEditActionByName = new Map(xlsxEditActions.map(value => [value.name, value]));

const tools = [
  {
    name: 'docx_inspect',
    description: 'Inspect a DOCX document and return a unified structural report including placeholders, comments, anchors, tables, fields, and formatting metrics.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string', description: 'Absolute or relative path to a .docx file.' } },
      required: ['input'],
    },
  },
  {
    name: 'docx_inspect_tables',
    description: 'Inspect DOCX body tables with row, cell, merge, paragraph alignment, run font, color, underline, and text-fill details.',
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
    name: 'docx_validate',
    description: 'Validate a DOCX package against the current Microsoft 365 OpenXML contract.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string' } },
      required: ['input'],
      additionalProperties: false,
    },
  },
  {
    name: 'docx_validate_font_policy',
    description: 'Independently validate current DOCX body and table runs against an explicit font policy.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        policy: {
          type: 'object',
          properties: {
            schema: { type: 'string' },
            body: { type: 'object' },
            table: { type: 'object' },
          },
          required: ['schema', 'body', 'table'],
          additionalProperties: false,
        },
      },
      required: ['input', 'policy'],
      additionalProperties: false,
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
  ...editToolDefinitions(docxEditActions),
  {
    name: 'xlsx_inspect',
    description: 'Inspect an XLSX workbook and return sheet-level metrics, used ranges, formula counts, and merged ranges.',
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
        resolveMergedCells: { type: 'boolean', description: 'Resolve merged cells to project values' }
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
  ...editToolDefinitions(xlsxEditActions),
  {
    name: 'xlsx_validate',
    description: 'Validate an XLSX workbook package and return Open XML validation evidence.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string', description: 'Absolute or relative path to a .xlsx file.' } },
      required: ['input'],
    },
  },
  {
    name: 'pptx_inspect',
    description: 'Inspect a PPTX file and return detailed master, layout, slide, shape, transform, paragraph, and run-format evidence.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string' } },
      required: ['input'],
    },
  },
  {
    name: 'pptx_export_json',
    description: 'Export PPTX slide text and placeholder hints as structured JSON.',
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
    name: 'pptx_fill_template',
    description: 'Fill PPTX text placeholders using a data object or JSON data file.',
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
    name: 'pptx_set_text_format',
    description: 'Set text-run formatting for current PPTX objects. A call batches only this one action kind.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
        changes: {
          type: 'array',
          minItems: 1,
          items: {
            type: 'object',
            properties: {
              slideNumber: { type: 'integer', minimum: 1 },
              shapeId: { type: 'integer', minimum: 1 },
              runIndex: integer,
              fontFamily: string,
              fontSize: { type: 'number', exclusiveMinimum: 0 },
              color: string,
              bold: boolean,
              paragraphAlignment: string,
            },
            required: ['slideNumber', 'shapeId', 'runIndex'],
            additionalProperties: false,
          },
        },
      },
      required: ['input', 'output', 'changes'],
      additionalProperties: false,
    },
  },
  {
    name: 'pptx_apply_template',
    description: 'Bind current slides to a selected master and layouts from a target PPTX template while preserving source business content.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        template: { type: 'string' },
        output: { type: 'string' },
        targetMasterPath: string,
        slides: {
          type: 'array',
          minItems: 1,
          items: {
            type: 'object',
            properties: {
              slideNumber: { type: 'integer', minimum: 1 },
              targetLayoutPath: string,
              contentBounds: {
                type: 'object',
                properties: { x: integer, y: integer, cx: { type: 'integer', minimum: 1 }, cy: { type: 'integer', minimum: 1 } },
                required: ['x', 'y', 'cx', 'cy'],
                additionalProperties: false,
              },
              contentShapeIds: { type: 'array', minItems: 1, uniqueItems: true, items: { type: 'integer', minimum: 1 } },
            },
            required: ['slideNumber', 'targetLayoutPath'],
            additionalProperties: false,
          },
        },
      },
      required: ['input', 'template', 'output', 'targetMasterPath', 'slides'],
      additionalProperties: false,
    },
  },
  {
    name: 'pptx_validate',
    description: 'Validate a PPTX package against the current Microsoft 365 OpenXML contract.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string' } },
      required: ['input'],
      additionalProperties: false,
    },
  },
];

async function callTool(name, args) {
  if (docxEditActionByName.has(name)) {
    return createToolResult(await fixedEdit(docxCandidates, docxEditActionByName.get(name), args));
  }
  if (xlsxEditActionByName.has(name)) {
    return createToolResult(await fixedEdit(xlsxCandidates, xlsxEditActionByName.get(name), args));
  }
  switch (name) {
    case 'docx_inspect':
      return createToolResult(await docxInspect(args));
    case 'docx_inspect_tables':
      return createToolResult(await docxInspectTables(args));
    case 'docx_compare':
      return createToolResult(await docxCompare(args));
    case 'docx_validate_template_transform':
      return createToolResult(await docxValidateTemplateTransform(args));
    case 'docx_validate':
      return createToolResult(await docxValidate(args));
    case 'docx_validate_font_policy':
      return createToolResult(await docxValidateFontPolicy(args));
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
    case 'xlsx_validate':
      return createToolResult(await xlsxValidate(args));
    case 'pptx_inspect':
      return createToolResult(await pptxInspect(args));
    case 'pptx_export_json':
      return createToolResult(await pptxExportJson(args));
    case 'pptx_fill_template':
      return createToolResult(await pptxFillTemplate(args));
    case 'pptx_set_text_format':
      return createToolResult(await pptxSetTextFormat(args));
    case 'pptx_apply_template':
      return createToolResult(await pptxApplyTemplate(args));
    case 'pptx_validate':
      return createToolResult(await pptxValidate(args));
    default:
      throw Object.assign(new Error(`Unknown tool: ${name}`), { code: -32601 });
  }
}

async function fixedEdit(candidates, actionDefinition, args) {
  const input = requireString(args.input, 'input');
  const output = requireString(args.output, 'output');
  if (actionDefinition.batch && (!Array.isArray(args.changes) || args.changes.length === 0)) {
    throw Object.assign(new Error('changes must be a non-empty array'), { code: -32602 });
  }
  const operations = actionDefinition.batch
    ? args.changes.map(change => ({ ...change, type: actionDefinition.operationType }))
    : [{ type: actionDefinition.operationType }];
  return withTempJsonFile({ operations }, async operationsPath => {
    const result = await runCandidateChain(candidates, ['edit', input, operationsPath, output]);
    return { tool: actionDefinition.name, runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  });
}

async function docxInspect(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(docxCandidates, ['inspect', input, '--json']);
  return { tool: 'docx_inspect', runtime: commandRuntime(result), report: result.json };
}

async function docxInspectTables(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(docxCandidates, ['inspect-tables', input, '--json']);
  return { tool: 'docx_inspect_tables', runtime: commandRuntime(result), report: result.json };
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

async function docxValidate(args) {
  const input = requireString(args.input, 'input');
  const result = await runValidationCandidateChain(docxCandidates, ['validate-openxml', input]);
  return { tool: 'docx_validate', runtime: commandRuntime(result), result: result.json };
}

async function docxValidateFontPolicy(args) {
  const input = requireString(args.input, 'input');
  if (!args.policy || typeof args.policy !== 'object' || Array.isArray(args.policy)) {
    throw Object.assign(new Error('policy must be an object'), { code: -32602 });
  }
  return withTempJsonFile(args.policy, async policyPath => {
    const result = await runValidationCandidateChain(docxCandidates, ['validate-font-policy', input, policyPath]);
    return { tool: 'docx_validate_font_policy', runtime: commandRuntime(result), result: result.json };
  });
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
  const cmdArgs = ['export-json', input];
  if (args.resolveMergedCells) {
    cmdArgs.push('--resolve-merged-cells');
  }
  if (args.output) {
    const output = requireString(args.output, 'output');
    cmdArgs.push(output);
    const result = await runCandidateChain(xlsxCandidates, cmdArgs);
    return { tool: 'xlsx_export_json', runtime: commandRuntime(result), outputPath: output, workbook: await maybeReadJson(output) };
  }
  const result = await runCandidateChain(xlsxCandidates, cmdArgs);
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

async function xlsxValidate(args) {
  const input = requireString(args.input, 'input');
  const result = await runValidationCandidateChain(xlsxCandidates, ['validate', input]);
  return { tool: 'xlsx_validate', runtime: commandRuntime(result), result: result.json };
}

async function pptxInspect(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(pptxCandidates, ['inspect', input, '--json']);
  return { tool: 'pptx_inspect', runtime: commandRuntime(result), report: result.json };
}

async function pptxExportJson(args) {
  const input = requireString(args.input, 'input');
  if (args.output) {
    const output = requireString(args.output, 'output');
    const result = await runCandidateChain(pptxCandidates, ['export-json', input, output]);
    return { tool: 'pptx_export_json', runtime: commandRuntime(result), outputPath: output, document: await maybeReadJson(output) };
  }
  const result = await runCandidateChain(pptxCandidates, ['export-json', input]);
  return { tool: 'pptx_export_json', runtime: commandRuntime(result), document: JSON.parse(result.stdout) };
}

async function pptxFillTemplate(args) {
  const template = requireString(args.template, 'template');
  const output = requireString(args.output, 'output');
  if (args.dataPath) {
    const dataPath = requireString(args.dataPath, 'dataPath');
    const result = await runCandidateChain(pptxCandidates, ['fill-template', template, dataPath, output]);
    return { tool: 'pptx_fill_template', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  }
  if (args.data === undefined) {
    throw Object.assign(new Error('data or dataPath is required'), { code: -32602 });
  }
  return withTempJsonFile(args.data, async dataPath => {
    const result = await runCandidateChain(pptxCandidates, ['fill-template', template, dataPath, output]);
    return { tool: 'pptx_fill_template', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  });
}

async function pptxSetTextFormat(args) {
  const input = requireString(args.input, 'input');
  const output = requireString(args.output, 'output');
  if (!Array.isArray(args.changes) || args.changes.length === 0) {
    throw Object.assign(new Error('changes must be a non-empty array'), { code: -32602 });
  }
  return withTempJsonFile({ operations: args.changes }, async planPath => {
    const result = await runCandidateChain(pptxCandidates, ['apply-format-edits', input, planPath, output]);
    return { tool: 'pptx_set_text_format', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  });
}

async function pptxApplyTemplate(args) {
  const input = requireString(args.input, 'input');
  const template = requireString(args.template, 'template');
  const output = requireString(args.output, 'output');
  const targetMasterPath = requireString(args.targetMasterPath, 'targetMasterPath');
  if (!Array.isArray(args.slides) || args.slides.length === 0) {
    throw Object.assign(new Error('slides must be a non-empty array'), { code: -32602 });
  }
  return withTempJsonFile({ targetMasterPath, slides: args.slides }, async planPath => {
    const result = await runCandidateChain(pptxCandidates, ['apply-template', input, template, planPath, output]);
    return { tool: 'pptx_apply_template', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  });
}

async function pptxValidate(args) {
  const input = requireString(args.input, 'input');
  const result = await runValidationCandidateChain(pptxCandidates, ['validate', input]);
  return { tool: 'pptx_validate', runtime: commandRuntime(result), result: result.json };
}
function commandRuntime(result) {
  return {
    command: result.command,
    cwd: result.cwd || path.dirname(result.command),
  };
}

await new McpStdioServer({ name: 'tiwater-office', version: '0.1.0', tools, callTool }).start();


async function runValidationCandidateChain(candidates, args) {
  const errors = [];
  for (const candidate of candidates) {
    try {
      const result = await runValidationCommand(candidate, args);
      const text = result.stdout.trim();
      if (!text) return { ...result, json: null };
      try {
        return { ...result, json: JSON.parse(text) };
      } catch {
        if (result.code !== 0) {
          errors.push(`${candidate.command}: validate did not return JSON`);
          continue;
        }
        throw new Error(`Expected JSON output but received: ${text.slice(0, 300)}${text.length > 300 ? '…' : ''}`);
      }
    } catch (error) {
      if (error?.code === 'ENOENT') {
        errors.push(`${candidate.command}: not found`);
        continue;
      }
      throw error;
    }
  }
  throw new Error(`No runnable command candidate succeeded. ${errors.join('; ')}`);
}

async function runValidationCommand(candidate, args) {
  const env = { ...process.env, ...(candidate.env || {}) };
  const cwd = candidate.cwd || resolveRepoPath();
  const commandArgs = [...(candidate.argsPrefix || []), ...args];

  return await new Promise((resolve, reject) => {
    const child = spawn(candidate.command, commandArgs, { cwd, env, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';

    child.stdout.on('data', chunk => {
      stdout += chunk.toString();
    });

    child.stderr.on('data', chunk => {
      stderr += chunk.toString();
    });

    child.on('error', reject);
    child.on('close', code => {
      if (code === 0 || code === 1) {
        resolve({ code, stdout, stderr, command: candidate.command, args: commandArgs });
        return;
      }
      reject(new Error(`${candidate.command} ${commandArgs.join(' ')} failed with exit code ${code}\n${stderr || stdout}`));
    });
  });
}
