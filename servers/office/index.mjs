#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { McpServer } from '@modelcontextprotocol/server';
import { serveStdio } from '@modelcontextprotocol/server/stdio';
import * as z from 'zod/v4';
import {
  commandCandidate,
  createToolResult,
  requireString,
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

const pathInput = z.string().trim().min(1);
const migrationAction = z.enum([
  'place-content',
  'keep-template-content',
  'keep-template-label',
  'select-template-option',
  'exclude-source',
  'review-source',
]);
const targetActions = new Set([
  'place-content',
  'keep-template-content',
  'keep-template-label',
  'select-template-option',
]);
const terminalActions = new Set(['exclude-source', 'review-source']);

const migrationChoiceInput = z.object({
  sourceChoiceId: z.string().trim().min(1),
  action: migrationAction,
  targetChoiceId: z.string().trim().min(1).optional(),
  cardinality: z.enum(['one', 'all']).optional(),
}).strict().superRefine((choice, context) => {
  if (targetActions.has(choice.action) && !choice.targetChoiceId) {
    context.addIssue({ code: 'custom', path: ['targetChoiceId'], message: `${choice.action} requires targetChoiceId` });
  }
  if (terminalActions.has(choice.action) && choice.targetChoiceId) {
    context.addIssue({ code: 'custom', path: ['targetChoiceId'], message: `${choice.action} forbids targetChoiceId` });
  }
  if (choice.cardinality === 'all' && !terminalActions.has(choice.action)) {
    context.addIssue({ code: 'custom', path: ['cardinality'], message: 'cardinality all is limited to terminal actions' });
  }
});

const templateCleanupInput = z.object({
  targetChoiceId: z.string().trim().min(1),
  scope: z.enum(['cell', 'row']),
}).strict();

const templateMigrationInput = z.object({
  source: pathInput.describe('Path to the current source DOCX.'),
  baseline: pathInput.describe('Path to the selected current baseline DOCX.'),
  output: pathInput.describe('Path to the migrated output DOCX.'),
  receiptOutput: pathInput.describe('New JSON receipt artifact path. Existing files are never overwritten.'),
  choices: z.array(migrationChoiceInput).describe('Exactly one business choice for every source id returned by docx_list_migration_choices.'),
  templateCleanup: z.array(templateCleanupInput).optional().describe('Optional baseline-owned placeholders or example rows to clear.'),
}).strict();

const runtimeIdentity = z.object({
  command: z.string(),
  cwd: z.string(),
}).strict();

const migrationChoiceOutput = z.object({
  id: z.string(),
  kind: z.string(),
  scope: z.string(),
  text: z.string().nullable(),
  count: z.number().int(),
  requiredCardinality: z.string().nullable(),
  context: z.record(z.string(), z.unknown()).nullable(),
  allowedActions: z.array(z.string()),
}).strict();

const migrationCatalog = z.object({
  schema: z.string(),
  pass: z.boolean(),
  sourceSha256: z.string(),
  baselineSha256: z.string(),
  sources: z.array(migrationChoiceOutput),
  targets: z.array(migrationChoiceOutput),
}).strict().superRefine((catalog, context) => {
  for (const key of ['sources', 'targets']) {
    const seen = new Set();
    for (const [index, choice] of catalog[key].entries()) {
      if (seen.has(choice.id)) {
        context.addIssue({ code: 'custom', path: [key, index, 'id'], message: `duplicate ${key} choice id: ${choice.id}` });
      }
      seen.add(choice.id);
    }
  }
});

const migrationReceipt = z.object({
  schema: z.string(),
  toolVersion: z.string(),
  status: z.enum(['pass', 'review-required', 'failed']),
  pass: z.boolean(),
  reviewRequired: z.boolean(),
  outputVerified: z.boolean(),
  output: z.string().nullable(),
  plan: z.string().nullable(),
  failures: z.array(z.unknown()),
}).passthrough();

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

const migrationCatalogOutput = z.object({
  tool: z.literal('docx_list_migration_choices'),
  runtime: runtimeIdentity,
  artifact,
  summary: z.object({
    schema: z.string(),
    pass: z.boolean(),
    sourceSha256: z.string(),
    baselineSha256: z.string(),
    sourceCount: z.number().int().nonnegative(),
    targetCount: z.number().int().nonnegative(),
  }).strict(),
}).strict();

const migrationQueryPage = z.object({
  offset: z.number().int().nonnegative(),
  returned: z.number().int().nonnegative(),
  total: z.number().int().nonnegative(),
  hasMore: z.boolean(),
}).strict();

const migrationChoiceQueryInput = z.discriminatedUnion('view', [
  z.object({
    catalog: pathInput.describe('Path returned by docx_list_migration_choices.'),
    view: z.literal('sources'),
    offset: z.number().int().nonnegative().optional(),
    limit: z.number().int().min(1).max(10).optional(),
  }).strict(),
  z.object({
    catalog: pathInput.describe('Path returned by docx_list_migration_choices.'),
    view: z.literal('targets'),
    sourceChoiceId: z.string().trim().min(1),
    text: z.string().trim().min(1).optional().describe('Literal case-insensitive text to find in target text or visible context.'),
    kinds: z.array(z.string().trim().min(1)).min(1).optional(),
    scopes: z.array(z.string().trim().min(1)).min(1).optional(),
    offset: z.number().int().nonnegative().optional(),
    limit: z.number().int().min(1).max(10).optional(),
  }).strict(),
]);

const migrationChoiceQueryOutput = z.object({
  tool: z.literal('docx_query_migration_choices'),
  catalogSha256: z.string().regex(/^[0-9a-f]{64}$/),
  view: z.enum(['sources', 'targets']),
  source: migrationChoiceOutput.nullable(),
  items: z.array(migrationChoiceOutput),
  page: migrationQueryPage,
}).strict();

function migrationReceiptOutput(tool) {
  return z.object({
    tool: z.literal(tool),
    runtime: runtimeIdentity,
    artifact,
    summary: z.object({
      schema: z.string(),
      toolVersion: z.string(),
      status: z.enum(['pass', 'review-required', 'failed']),
      pass: z.boolean(),
      reviewRequired: z.boolean(),
      outputVerified: z.boolean(),
      output: z.string().nullable(),
      plan: z.string().nullable(),
      failureCount: z.number().int().nonnegative(),
    }).strict(),
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
    name: 'docx_list_migration_choices',
    description: 'Write every current source item that still needs a business choice and the selectable current baseline targets to a run-local JSON artifact. Returns only artifact metadata and counts; it does not recommend a choice.',
    inputSchema: z.object({
      source: pathInput.describe('Path to the current source DOCX.'),
      baseline: pathInput.describe('Path to the selected current baseline DOCX.'),
      output: pathInput.describe('New JSON artifact path. Existing files are never overwritten.'),
    }).strict(),
    outputSchema: migrationCatalogOutput,
    handler: docxListMigrationChoices,
  },
  {
    name: 'docx_query_migration_choices',
    description: 'Read one bounded page from a migration-choice catalog. List source choices, or inspect targets for one source using literal text, kind, and scope filters. This tool does not recommend or make a business choice.',
    inputSchema: migrationChoiceQueryInput,
    outputSchema: migrationChoiceQueryOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxQueryMigrationChoices,
  },
  {
    name: 'docx_migrate_template',
    description: 'Migrate a current DOCX into the selected baseline from one complete batch of business choices. Choices reference only opaque ids returned by docx_list_migration_choices; the tool derives all document values, coordinates, plans, and edits.',
    inputSchema: templateMigrationInput,
    outputSchema: migrationReceiptOutput('docx_migrate_template'),
    handler: docxMigrateTemplate,
  },
  {
    name: 'docx_verify_migration',
    description: 'Independently re-resolve the same business choices and verify a migrated DOCX against the current source and baseline. This does not trust the migration receipt.',
    inputSchema: templateMigrationInput,
    outputSchema: migrationReceiptOutput('docx_verify_migration'),
    handler: docxVerifyMigration,
  },
  {
    name: 'docx_compare',
    description: 'Compare two DOCX files and report package, metric, and style differences.',
    inputSchema: z.object({ baseline: pathInput, updated: pathInput }).strict(),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxCompare,
  },
  {
    name: 'docx_validate_template_transform',
    description: 'Validate whether a source DOCX template and target DOCX template are structurally compatible.',
    inputSchema: z.object({ sourceTemplate: pathInput, targetTemplate: pathInput }).strict(),
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxValidateTemplateTransform,
  },
  {
    name: 'docx_export_json',
    description: 'Export DOCX body content to a new JSON artifact without returning the full document through MCP.',
    inputSchema: artifactInput,
    outputSchema: artifactOutput('docx_export_json'),
    handler: docxExportJson,
  },
  {
    name: 'xlsx_inspect',
    description: 'Inspect an XLSX workbook and write one JSON observation containing workbook structure, exported values, formulas, styles, merged ranges, and conversion evidence.',
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
];

function buildServer() {
  const server = new McpServer(
    { name: 'tiwater-office', version: packageMetadata.version },
    {
      instructions: 'Use the Office tools for technical document observation. For template migration, list the current choices, select only allowed business actions, migrate once, and independently verify the result. Never invent document values, identities, coordinates, plans, or edit operations.',
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

async function docxListMigrationChoices(args) {
  const source = requireString(args.source, 'source');
  const baseline = requireString(args.baseline, 'baseline');
  const result = await runJsonCandidateChain(docxCandidates, ['list-template-migration-choices', source, baseline]);
  const catalog = migrationCatalog.parse(result.json);
  return {
    tool: 'docx_list_migration_choices',
    runtime: commandRuntime(result),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), catalog),
    summary: {
      schema: catalog.schema,
      pass: catalog.pass,
      sourceSha256: catalog.sourceSha256,
      baselineSha256: catalog.baselineSha256,
      sourceCount: catalog.sources.length,
      targetCount: catalog.targets.length,
    },
  };
}

async function docxQueryMigrationChoices(args) {
  const catalogPath = requireString(args.catalog, 'catalog');
  const bytes = await readFile(catalogPath);
  const catalog = migrationCatalog.parse(JSON.parse(bytes.toString('utf8')));
  const offset = args.offset ?? 0;
  const limit = args.limit ?? 10;
  let source = null;
  let matches;

  if (args.view === 'sources') {
    matches = catalog.sources;
  } else {
    source = catalog.sources.find(item => item.id === args.sourceChoiceId) ?? null;
    if (!source) {
      throw Object.assign(new Error(`Unknown sourceChoiceId: ${args.sourceChoiceId}`), { code: -32602 });
    }
    const kinds = args.kinds ? new Set(args.kinds) : null;
    const scopes = args.scopes ? new Set(args.scopes) : null;
    const textQuery = args.text?.toLocaleLowerCase();
    matches = catalog.targets.filter(item =>
      (!kinds || kinds.has(item.kind)) &&
      (!scopes || scopes.has(item.scope)) &&
      (!textQuery || migrationChoiceSearchText(item).includes(textQuery)));
  }

  const items = matches.slice(offset, offset + limit);
  return {
    tool: 'docx_query_migration_choices',
    catalogSha256: createHash('sha256').update(bytes).digest('hex'),
    view: args.view,
    source,
    items,
    page: {
      offset,
      returned: items.length,
      total: matches.length,
      hasMore: offset + items.length < matches.length,
    },
  };
}

function migrationChoiceSearchText(choice) {
  return collectVisibleStrings({ text: choice.text, context: choice.context })
    .join('\n')
    .toLocaleLowerCase();
}

function collectVisibleStrings(value, fieldName = '') {
  if (typeof value === 'string') return /text/i.test(fieldName) ? [value] : [];
  if (Array.isArray(value)) return value.flatMap(item => collectVisibleStrings(item, fieldName));
  if (value && typeof value === 'object') {
    return Object.entries(value).flatMap(([key, item]) => collectVisibleStrings(item, key));
  }
  return [];
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
    const receipt = migrationReceipt.parse(result.json);
    return {
      tool,
      runtime: commandRuntime(result),
      artifact: await writeJsonArtifact(requireString(args.receiptOutput, 'receiptOutput'), receipt),
      summary: {
        schema: receipt.schema,
        toolVersion: receipt.toolVersion,
        status: receipt.status,
        pass: receipt.pass,
        reviewRequired: receipt.reviewRequired,
        outputVerified: receipt.outputVerified,
        output: receipt.output,
        plan: receipt.plan,
        failureCount: receipt.failures.length,
      },
    };
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

async function docxExportJson(args) {
  const input = requireString(args.input, 'input');
  const result = await runJsonCandidateChain(docxCandidates, ['export-json', input]);
  return {
    tool: 'docx_export_json',
    runtime: commandRuntime(result),
    artifact: await writeJsonArtifact(requireString(args.output, 'output'), result.json),
  };
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
  const result = await runXlsxValidateCandidateChain(['validate', input]);
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

function commandRuntime(result) {
  return {
    command: result.command,
    cwd: result.cwd || path.dirname(result.command),
  };
}

serveStdio(buildServer);

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
  const cwd = candidate.cwd || process.cwd();
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
        resolve({ code, stdout, stderr, command: candidate.command, args: commandArgs, cwd });
        return;
      }
      reject(new Error(`${candidate.command} ${commandArgs.join(' ')} failed with exit code ${code}\n${stderr || stdout}`));
    });
  });
}
