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

const docxCandidates = [
  commandCandidate('tiwater-docx'),
];

const xlsxCandidates = [
  commandCandidate('tiwater-xlsx'),
];

const pptxCandidates = [
  commandCandidate('tiwater-pptx'),
];

function templateMigrationInputSchema() {
  return {
    type: 'object',
    properties: {
      source: { type: 'string', description: 'Path to the current source DOCX.' },
      baseline: { type: 'string', description: 'Path to the selected current baseline DOCX.' },
      output: { type: 'string', description: 'Path to the migrated output DOCX.' },
      choices: {
        type: 'array',
        description: 'Exactly one business choice for every source id returned by docx_list_migration_choices.',
        items: {
          type: 'object',
          properties: {
            sourceChoiceId: { type: 'string' },
            action: {
              type: 'string',
              enum: ['place-content', 'keep-template-content', 'keep-template-label', 'select-template-option', 'exclude-source', 'review-source'],
            },
            targetChoiceId: { type: 'string', description: 'Required only when the selected action uses a baseline target.' },
            cardinality: { type: 'string', enum: ['one', 'all'] },
          },
          required: ['sourceChoiceId', 'action'],
          additionalProperties: false,
        },
      },
      templateCleanup: {
        type: 'array',
        description: 'Optional baseline-owned placeholders or example rows to clear.',
        items: {
          type: 'object',
          properties: {
            targetChoiceId: { type: 'string' },
            scope: { type: 'string', enum: ['cell', 'row'] },
          },
          required: ['targetChoiceId', 'scope'],
          additionalProperties: false,
        },
      },
    },
    required: ['source', 'baseline', 'output', 'choices'],
    additionalProperties: false,
  };
}

const runtimeIdentitySchema = {
  type: 'object',
  properties: {
    command: { type: 'string' },
    cwd: { type: 'string' },
  },
  required: ['command', 'cwd'],
  additionalProperties: false,
};

const migrationChoiceSchema = {
  type: 'object',
  properties: {
    id: { type: 'string' },
    kind: { type: 'string' },
    scope: { type: 'string' },
    text: { type: ['string', 'null'] },
    count: { type: 'integer' },
    requiredCardinality: { type: ['string', 'null'] },
    context: { type: ['object', 'null'] },
    allowedActions: { type: 'array', items: { type: 'string' } },
  },
  required: ['id', 'kind', 'scope', 'text', 'count', 'requiredCardinality', 'context', 'allowedActions'],
  additionalProperties: false,
};

const migrationCatalogOutputSchema = {
  type: 'object',
  properties: {
    tool: { const: 'docx_list_migration_choices' },
    runtime: runtimeIdentitySchema,
    catalog: {
      type: 'object',
      properties: {
        schema: { type: 'string' },
        pass: { type: 'boolean' },
        sourceSha256: { type: 'string' },
        baselineSha256: { type: 'string' },
        sources: { type: 'array', items: migrationChoiceSchema },
        targets: { type: 'array', items: migrationChoiceSchema },
      },
      required: ['schema', 'pass', 'sourceSha256', 'baselineSha256', 'sources', 'targets'],
      additionalProperties: false,
    },
  },
  required: ['tool', 'runtime', 'catalog'],
  additionalProperties: false,
};

function migrationReceiptOutputSchema(tool) {
  return {
    type: 'object',
    properties: {
      tool: { const: tool },
      runtime: runtimeIdentitySchema,
      receipt: {
        type: 'object',
        properties: {
          schema: { type: 'string' },
          toolVersion: { type: 'string' },
          status: { type: 'string', enum: ['pass', 'review-required', 'failed'] },
          pass: { type: 'boolean' },
          reviewRequired: { type: 'boolean' },
          outputVerified: { type: 'boolean' },
          output: { type: ['string', 'null'] },
          plan: { type: ['string', 'null'] },
          failures: { type: 'array', items: { type: 'object' } },
        },
        required: ['schema', 'toolVersion', 'status', 'pass', 'reviewRequired', 'outputVerified', 'output', 'plan', 'failures'],
        additionalProperties: true,
      },
    },
    required: ['tool', 'runtime', 'receipt'],
    additionalProperties: false,
  };
}

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
    name: 'docx_list_migration_choices',
    description: 'List every current source item that still needs a business choice and the selectable current baseline targets. Returns opaque ids and context; it does not recommend a choice.',
    inputSchema: {
      type: 'object',
      properties: {
        source: { type: 'string', description: 'Path to the current source DOCX.' },
        baseline: { type: 'string', description: 'Path to the selected current baseline DOCX.' },
      },
      required: ['source', 'baseline'],
      additionalProperties: false,
    },
    outputSchema: migrationCatalogOutputSchema,
  },
  {
    name: 'docx_migrate_template',
    description: 'Migrate a current DOCX into the selected baseline from one complete batch of business choices. Choices reference only opaque ids returned by docx_list_migration_choices; the tool derives all document values, coordinates, plans, and edits.',
    inputSchema: templateMigrationInputSchema(),
    outputSchema: migrationReceiptOutputSchema('docx_migrate_template'),
  },
  {
    name: 'docx_verify_migration',
    description: 'Independently re-resolve the same business choices and verify a migrated DOCX against the current source and baseline. This does not trust the migration receipt.',
    inputSchema: templateMigrationInputSchema(),
    outputSchema: migrationReceiptOutputSchema('docx_verify_migration'),
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
];

async function callTool(name, args) {
  switch (name) {
    case 'docx_inspect':
      return createToolResult(await docxInspect(args));
    case 'docx_inspect_tables':
      return createToolResult(await docxInspectTables(args));
    case 'docx_list_migration_choices':
      return createToolResult(await docxListMigrationChoices(args));
    case 'docx_migrate_template':
      return createToolResult(await docxMigrateTemplate(args));
    case 'docx_verify_migration':
      return createToolResult(await docxVerifyMigration(args));
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
    case 'xlsx_inspect':
      return createToolResult(await xlsxInspect(args));
    case 'xlsx_export_json':
      return createToolResult(await xlsxExportJson(args));
    case 'xlsx_validate':
      return createToolResult(await xlsxValidate(args));
    case 'pptx_inspect':
      return createToolResult(await pptxInspect(args));
    case 'pptx_inspect_detail':
      return createToolResult(await pptxInspectDetail(args));
    case 'pptx_export_json':
      return createToolResult(await pptxExportJson(args));
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

async function docxListMigrationChoices(args) {
  const source = requireString(args.source, 'source');
  const baseline = requireString(args.baseline, 'baseline');
  const result = await runJsonCandidateChain(docxCandidates, ['list-template-migration-choices', source, baseline]);
  return { tool: 'docx_list_migration_choices', runtime: commandRuntime(result), catalog: result.json };
}

async function docxMigrateTemplate(args) {
  return runTemplateMigrationCommand('docx_migrate_template', 'migrate-template', args);
}

async function docxVerifyMigration(args) {
  return runTemplateMigrationCommand('docx_verify_migration', 'verify-template-migration', args);
}

async function runTemplateMigrationCommand(tool, command, args) {
  const source = requireString(args.source, 'source');
  const baseline = requireString(args.baseline, 'baseline');
  const output = requireString(args.output, 'output');
  if (!Array.isArray(args.choices)) {
    throw Object.assign(new Error('choices must be an array'), { code: -32602 });
  }
  const payload = {
    schema: 'tiwater.docx.template-migration-business-choices/v1',
    choices: args.choices,
    ...(Array.isArray(args.templateCleanup) ? { templateCleanup: args.templateCleanup } : {}),
  };
  return withTempJsonFile(payload, async choicesPath => {
    const result = await runJsonCandidateChain(
      docxCandidates,
      [command, source, baseline, choicesPath, output],
      { allowedExitCodes: [0, 1] });
    return { tool, runtime: commandRuntime(result), receipt: result.json };
  });
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

function commandRuntime(result) {
  return {
    command: result.command,
    cwd: result.cwd || path.dirname(result.command),
  };
}

await new McpStdioServer({ name: 'tiwater-office', version: '0.2.0', tools, callTool }).start();

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
