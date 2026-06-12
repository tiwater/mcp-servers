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
    name: 'docx_edit',
    description: 'Apply explicit edit operations to a DOCX document, including anchored text replacement, paragraph/cell edits, rich text table cells, cell merging, comment deletion, and field refresh.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
        edits: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              type: { type: 'string', enum: ['replaceAnchoredText', 'replaceParagraphText', 'replaceBodyText', 'replaceAllHeaderParagraphText', 'replaceHeaderParagraphText', 'replaceHeaderText', 'replaceTableCellText', 'replaceTableCellRichText', 'replaceTable', 'insertTableRows', 'deleteTableRows', 'replaceTableRows', 'insertTableColumns', 'setTableWidth', 'setTableCellAlignment', 'setTableCellNoWrap', 'setTableCellFontSize', 'setTableRowHeight', 'mergeTableCells', 'unmergeTableRowHorizontalCells', 'unmergeTableColumnVerticalCells', 'fillTableSemantically', 'deleteComment', 'deleteComments', 'markFieldsDirty', 'sanitizeFields', 'freezeFields'] },
              commentId: { type: 'string' },
              text: { type: 'string' },
              richText: {
                type: 'array',
                items: {
                  type: 'object',
                  properties: {
                    text: { type: 'string' },
                    color: { type: 'string' },
                    underline: { type: 'boolean' },
                    bold: { type: 'boolean' }
                  },
                  required: ['text']
                }
              },
              findText: { type: 'string' },
              paragraphIndex: { type: 'integer' },
              headerIndex: { type: 'integer' },
              tableIndex: { type: 'integer' },
              rowIndex: { type: 'integer' },
              cellIndex: { type: 'integer' },
              commentIds: { type: 'array', items: { type: 'string' } },
              startCellIndex: { type: 'integer' },
              endCellIndex: { type: 'integer' },
              startRowIndex: { type: 'integer' },
              endRowIndex: { type: 'integer' },
              columnIndex: { type: 'integer' },
              columnCount: { type: 'integer' },
              templateColumnIndex: { type: 'integer' },
              templateRowIndex: { type: 'integer' },
              width: { type: 'string' },
              widthType: { type: 'string', enum: ['pct', 'dxa', 'auto', 'nil'] },
              alignment: { type: 'string', enum: ['left', 'center', 'right', 'both'] },
              noWrap: { type: 'boolean', description: 'When true or omitted, set Word w:noWrap on the target table cell; false removes it.' },
              fontSize: { type: 'string', description: 'OpenXML half-points such as 18, or point size such as 9pt.' },
              height: { type: 'string', description: 'Table row height in twips.' },
              heightRule: { type: 'string', enum: ['atLeast', 'at-least', 'at_least', 'exact', 'auto'] },
              rows: {
                type: 'array',
                items: {
                  type: 'array',
                  items: {
                    type: 'object',
                    properties: {
                      text: { type: 'string' },
                      gridSpan: { type: 'integer' },
                      vMerge: { type: 'string', enum: ['restart', 'continue'] },
                      bold: { type: 'boolean' },
                      header: { type: 'boolean' },
                      shading: { type: 'string' },
                      alignment: { type: 'string', enum: ['left', 'center', 'right'] },
                      richText: {
                        type: 'array',
                        items: {
                          type: 'object',
                          properties: {
                            text: { type: 'string' },
                            color: { type: 'string' },
                            underline: { type: 'boolean' },
                            bold: { type: 'boolean' }
                          },
                          required: ['text']
                        }
                      }
                    }
                  }
                }
               },
              cells: {
                type: 'array',
                items: {
                  type: 'object',
                  properties: {
                    rowPatterns: { type: 'array', items: { type: 'string' } },
                    colPatterns: { type: 'array', items: { type: 'string' } },
                    text: { type: 'string' }
                  },
                  required: ['rowPatterns', 'colPatterns', 'text']
                }
              }
            },
            required: ['type']
          }
        },
        editsPath: { type: 'string' }
      },
      required: ['input', 'output'],
    },
  },
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
  {
    name: 'xlsx_edit',
    description: 'Apply explicit edit operations to an XLSX workbook, including single-cell writes, range writes, structural row operations, and anchored section expansion for fixed-layout sheets.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
        edits: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              type: { type: 'string' },
              sheet: { type: 'string' },
              cell: { type: 'string' },
              value: { type: 'string' },
              valueType: { type: 'string' },
              bold: { type: 'boolean' },
              startCell: { type: 'string' },
              values: { type: 'array', items: { type: 'array', items: { type: 'string' } } },
              startRow: { type: 'integer' },
              count: { type: 'integer' },
              sourceRow: { type: 'integer' },
              targetRow: { type: 'integer' },
              translateFormulas: { type: 'boolean' },
              anchorText: { type: 'string' },
              exampleRows: { type: 'integer' },
              targetRows: { type: 'integer' },
              preserveStyle: { type: 'boolean' },
              preserveFormulas: { type: 'boolean' },
              preserveMergedRanges: { type: 'boolean' }
            },
            required: ['type']
          }
        },
        editsPath: { type: 'string' }
      },
      required: ['input', 'output'],
    },
  },
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
    description: 'Inspect a PPTX file and return slide metrics and discovered placeholders.',
    inputSchema: {
      type: 'object',
      properties: { input: { type: 'string' } },
      required: ['input'],
    },
  },
  {
    name: 'pptx_inspect_detail',
    description: 'Inspect a PPTX file and return detailed slide, shape, transform, paragraph, and run-format evidence.',
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
    name: 'pptx_apply_format_edits',
    description: 'Copy a PPTX and apply targeted run-format edits from a data object or JSON edit plan file.',
    inputSchema: {
      type: 'object',
      properties: {
        input: { type: 'string' },
        output: { type: 'string' },
        plan: { type: 'object' },
        planPath: { type: 'string' },
      },
      required: ['input', 'output'],
    },
  },
];

async function callTool(name, args) {
  switch (name) {
    case 'docx_inspect':
      return createToolResult(await docxInspect(args));
    case 'docx_inspect_tables':
      return createToolResult(await docxInspectTables(args));
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
    case 'docx_edit':
      return createToolResult(await docxEdit(args));
    case 'xlsx_inspect':
      return createToolResult(await xlsxInspect(args));
    case 'xlsx_export_json':
      return createToolResult(await xlsxExportJson(args));
    case 'xlsx_fill_template':
      return createToolResult(await xlsxFillTemplate(args));
    case 'xlsx_edit':
      return createToolResult(await xlsxEdit(args));
    case 'xlsx_validate':
      return createToolResult(await xlsxValidate(args));
    case 'pptx_inspect':
      return createToolResult(await pptxInspect(args));
    case 'pptx_inspect_detail':
      return createToolResult(await pptxInspectDetail(args));
    case 'pptx_export_json':
      return createToolResult(await pptxExportJson(args));
    case 'pptx_fill_template':
      return createToolResult(await pptxFillTemplate(args));
    case 'pptx_apply_format_edits':
      return createToolResult(await pptxApplyFormatEdits(args));
    default:
      throw Object.assign(new Error(`Unknown tool: ${name}`), { code: -32601 });
  }
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

async function docxEdit(args) {
  const input = requireString(args.input, 'input');
  const output = requireString(args.output, 'output');
  if (args.editsPath) {
    const editsPath = requireString(args.editsPath, 'editsPath');
    const result = await runCandidateChain(docxCandidates, ['edit', input, editsPath, output]);
    return { tool: 'docx_edit', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  }
  if (!Array.isArray(args.edits)) {
    throw Object.assign(new Error('edits or editsPath is required'), { code: -32602 });
  }
  return withTempJsonFile({ operations: args.edits }, async editsPath => {
    const result = await runCandidateChain(docxCandidates, ['edit', input, editsPath, output]);
    return { tool: 'docx_edit', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
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
  const result = await runXlsxValidateCandidateChain(['validate', input]);
  return { tool: 'xlsx_validate', runtime: commandRuntime(result), result: result.json };
}

async function pptxInspect(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(pptxCandidates, ['inspect', input, '--json']);
  return { tool: 'pptx_inspect', runtime: commandRuntime(result), report: result.json };
}

async function pptxInspectDetail(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(pptxCandidates, ['inspect', input, '--json', '--detail']);
  return { tool: 'pptx_inspect_detail', runtime: commandRuntime(result), report: result.json };
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

async function pptxApplyFormatEdits(args) {
  const input = requireString(args.input, 'input');
  const output = requireString(args.output, 'output');
  if (args.planPath) {
    const planPath = requireString(args.planPath, 'planPath');
    const result = await runCandidateChain(pptxCandidates, ['apply-format-edits', input, planPath, output]);
    return { tool: 'pptx_apply_format_edits', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  }
  if (args.plan === undefined) {
    throw Object.assign(new Error('plan or planPath is required'), { code: -32602 });
  }
  return withTempJsonFile(args.plan, async planPath => {
    const result = await runCandidateChain(pptxCandidates, ['apply-format-edits', input, planPath, output]);
    return { tool: 'pptx_apply_format_edits', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  });
}
function commandRuntime(result) {
  return {
    command: result.command,
    cwd: result.cwd || path.dirname(result.command),
  };
}

await new McpStdioServer({ name: 'tiwater-office', version: '0.1.0', tools, callTool }).start();


async function xlsxEdit(args) {
  const input = requireString(args.input, 'input');
  const output = requireString(args.output, 'output');
  if (args.editsPath) {
    const editsPath = requireString(args.editsPath, 'editsPath');
    const result = await runCandidateChain(xlsxCandidates, ['edit', input, editsPath, output]);
    return { tool: 'xlsx_edit', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  }
  if (!Array.isArray(args.edits)) {
    throw Object.assign(new Error('edits or editsPath is required'), { code: -32602 });
  }
  return withTempJsonFile({ operations: args.edits }, async editsPath => {
    const result = await runCandidateChain(xlsxCandidates, ['edit', input, editsPath, output]);
    return { tool: 'xlsx_edit', runtime: commandRuntime(result), outputPath: output, result: JSON.parse(result.stdout) };
  });
}

async function runXlsxValidateCandidateChain(args) {
  const errors = [];
  for (const candidate of xlsxCandidates) {
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
