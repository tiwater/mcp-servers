#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { createReadStream } from 'node:fs';
import { mkdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { isDeepStrictEqual } from 'node:util';
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

const convertCandidates = [
  commandCandidate('tiwater-convert', [], { cwd: invocationCwd }),
];

const pathInput = z.string().trim().min(1);
const migrationQueryTargetAction = z.enum([
  'place-content',
  'keep-template-label',
  'select-template-option',
]);
const targetActions = new Set(migrationQueryTargetAction.options);

const choiceReference = z.string().trim().regex(/^[ST][1-9][0-9]*-[0-9a-f]{8}$/);
const alternativeReference = z.string().trim().regex(/^S[1-9][0-9]*-[PLO][1-9][0-9]*-[0-9a-f]{8}$/);

const terminalMigrationChoiceInput = z.object({
  sourceRef: choiceReference,
  action: z.enum(['exclude-source', 'review-source']),
}).strict();

const migrationChoiceInput = z.union([
  z.object({ alternativeRef: alternativeReference }).strict(),
  terminalMigrationChoiceInput,
]);

const templateCleanupInput = z.object({
  targetRef: choiceReference,
  scope: z.enum(['cell', 'row']),
}).strict();

const templateMigrationInput = z.object({
  source: pathInput.describe('Path to the current source DOCX.'),
  baseline: pathInput.describe('Path to the selected current baseline DOCX.'),
  output: pathInput.describe('Path to the migrated output DOCX.'),
  receiptOutput: pathInput.describe('New JSON receipt artifact path. Existing files are never overwritten.'),
  choices: z.array(migrationChoiceInput).describe('Exactly one business choice for every source ref. Targeted choices use an alternativeRef returned by the query tool. Terminal exclusions and genuine local review use a sourceRef.'),
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

const migrationPublicChoiceOutput = migrationChoiceOutput.omit({ id: true }).extend({
  ref: choiceReference,
}).strict();

const migrationAlternativeOutput = z.object({
  ref: alternativeReference,
  action: migrationQueryTargetAction,
  target: migrationPublicChoiceOutput,
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

const migrationTargetPage = z.object({
  schema: z.string(),
  pass: z.boolean(),
  sourceChoiceId: z.string().nullable(),
  branch: z.string(),
  offset: z.number().int().nonnegative(),
  limit: z.number().int().positive(),
  total: z.number().int().nonnegative(),
  targets: z.array(migrationChoiceOutput),
}).strict();

const migrationQueryDocuments = {
  source: pathInput.describe('Path to the current source DOCX used by docx_list_migration_choices.'),
  baseline: pathInput.describe('Path to the same selected current baseline DOCX used by docx_list_migration_choices.'),
};

const boundedOffset = z.number().int().nonnegative().optional().describe('Zero-based result offset. Defaults to 0.');
const boundedLimit = z.number().int().min(1).max(10).optional().describe('Maximum results to return. Defaults to 10 and cannot exceed 10.');

const migrationChoiceQueryInput = z.union([
  z.object({
    ...migrationQueryDocuments,
    view: z.literal('sources'),
    offset: boundedOffset,
    limit: boundedLimit,
  }).strict(),
  z.object({
    ...migrationQueryDocuments,
    view: z.literal('targets'),
    sourceRef: choiceReference.describe('Short source reference returned by the sources view.'),
    action: migrationQueryTargetAction.optional().describe('Optional action filter. Omit it to inspect every independently valid action-and-target alternative for the source.'),
    text: z.string().trim().min(1).optional().describe('Optional literal case-insensitive text to find in target visible text or context.'),
    offset: boundedOffset,
    limit: boundedLimit,
  }).strict(),
  z.object({
    ...migrationQueryDocuments,
    view: z.literal('cleanup'),
    text: z.string().trim().min(1).optional().describe('Optional literal case-insensitive text to find in cleanup target visible text or context.'),
    offset: boundedOffset,
    limit: boundedLimit,
  }).strict(),
]);

const migrationChoiceQueryOutput = z.object({
  tool: z.literal('docx_query_migration_choices'),
  runtime: runtimeIdentity,
  sourceSha256: z.string(),
  baselineSha256: z.string(),
  view: z.enum(['sources', 'targets', 'cleanup']),
  action: migrationQueryTargetAction.nullable(),
  source: migrationPublicChoiceOutput.nullable(),
  items: z.array(z.union([migrationPublicChoiceOutput, migrationAlternativeOutput])),
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
    description: 'Write every current source item that still needs a business choice and the selectable current baseline targets to an opaque run-local evidence artifact. Use docx_query_migration_choices with the same source and baseline to inspect bounded alternatives; do not parse the artifact.',
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
    description: 'Query current template-migration alternatives without reading the catalog artifact. Source and cleanup results use short catalog-bound refs. Target results bind one provider-compatible action and target into a single alternativeRef; select that ref without reconstructing the pair. Action and text filters are optional. Result order helps discovery but does not make the business choice.',
    inputSchema: migrationChoiceQueryInput,
    outputSchema: migrationChoiceQueryOutput,
    annotations: { readOnlyHint: true, idempotentHint: true },
    handler: docxQueryMigrationChoices,
  },
  {
    name: 'docx_migrate_template',
    description: 'Migrate a current DOCX into the selected baseline from one complete batch of business choices. For targeted choices, submit one alternativeRef returned by docx_query_migration_choices. Exclusion and local review use a sourceRef plus their terminal action. The tool derives all document values, coordinates, plans, and edits.',
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
    name: 'docx_export_json',
    description: 'Export DOCX body content to a new JSON artifact without returning the full document through MCP.',
    inputSchema: artifactInput,
    outputSchema: artifactOutput('docx_export_json'),
    handler: docxExportJson,
  },
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
  const sourcePath = requireString(args.source, 'source');
  const baselinePath = requireString(args.baseline, 'baseline');
  const offset = args.offset ?? 0;
  const limit = args.limit ?? 10;
  const catalogResult = await runJsonCandidateChain(docxCandidates, ['list-template-migration-choices', sourcePath, baselinePath]);
  const catalog = migrationCatalog.parse(catalogResult.json);

  if (args.view === 'sources') {
    return migrationQueryResult({
      runtime: commandRuntime(catalogResult),
      catalog,
      view: 'sources',
      action: null,
      source: null,
      items: catalog.sources.slice(offset, offset + limit),
      offset,
      total: catalog.sources.length,
    });
  }

  let source = null;
  let action = null;
  if (args.view === 'cleanup') {
    const targetResult = await loadMigrationTargets({
      sourcePath,
      baselinePath,
      sourceChoiceId: '-',
      branch: 'baseline-clear',
      text: args.text,
    });
    const orderedTargets = validateAndOrderMigrationTargets(catalog, null, targetResult.targets);
    return migrationQueryResult({
      runtime: targetResult.runtime,
      catalog,
      view: 'cleanup',
      action: null,
      source: null,
      items: orderedTargets.slice(offset, offset + limit),
      offset,
      total: orderedTargets.length,
    });
  }

  const requestedSourceId = choiceIdFromRef(catalog, catalog.sources, args.sourceRef, 'S', 'source');
  source = catalog.sources.find(item => item.id === requestedSourceId) ?? null;
  if (!source) {
    throw Object.assign(new Error(`Unknown migration source: ${args.sourceRef}`), { code: -32602 });
  }
  action = args.action ?? null;
  const actions = action
    ? [action]
    : [...targetActions].filter(candidate => source.allowedActions.includes(candidate));
  const pages = await Promise.all(actions.map(async candidate => {
    if (!source.allowedActions.includes(candidate)) {
      throw Object.assign(new Error(`Action ${candidate} is not allowed for ${source.id}`), { code: -32602 });
    }
    const result = await loadMigrationTargets({
      sourcePath,
      baselinePath,
      sourceChoiceId: source.id,
      branch: migrationTargetBranch(candidate, source.kind),
      text: args.text,
    });
    return {
      action: candidate,
      runtime: result.runtime,
      targets: validateAndOrderMigrationTargets(catalog, source, result.targets),
    };
  }));
  const catalogOrder = new Map(catalog.targets.map((item, index) => [item.id, index]));
  const actionOrder = new Map([...targetActions].map((item, index) => [item, index]));
  const alternatives = pages.flatMap(page => page.targets.map(target => ({ action: page.action, target })))
    .sort((left, right) => {
      const relevance = compareRelevance(
        migrationChoiceRelevance(source, right.target),
        migrationChoiceRelevance(source, left.target));
      return relevance
        || catalogOrder.get(left.target.id) - catalogOrder.get(right.target.id)
        || actionOrder.get(left.action) - actionOrder.get(right.action);
    })
    .map(item => ({
      ref: migrationAlternativeRef(catalog, source, item.action, item.target),
      action: item.action,
      target: publicMigrationChoice(catalog, item.target, 'T'),
    }));
  return migrationQueryResult({
    runtime: pages[0]?.runtime ?? commandRuntime(catalogResult),
    catalog,
    view: 'targets',
    action,
    source,
    items: alternatives.slice(offset, offset + limit),
    offset,
    total: alternatives.length,
    alternatives: true,
  });
}

function validateAndOrderMigrationTargets(catalog, source, targets) {
  const catalogTargets = new Map(catalog.targets.map(item => [item.id, item]));
  for (const target of targets) {
    const current = catalogTargets.get(target.id);
    if (!current || !isDeepStrictEqual(current, target)) {
      throw new Error(`Migration target ${target.id} is not bound to the current catalog`);
    }
  }
  const catalogOrder = new Map(catalog.targets.map((item, index) => [item.id, index]));
  return source
    ? [...targets].sort((left, right) => {
      const relevance = compareRelevance(
        migrationChoiceRelevance(source, right),
        migrationChoiceRelevance(source, left));
      return relevance || catalogOrder.get(left.id) - catalogOrder.get(right.id);
    })
    : [...targets].sort((left, right) => catalogOrder.get(left.id) - catalogOrder.get(right.id));
}

async function loadMigrationTargets({ sourcePath, baselinePath, sourceChoiceId, branch, text }) {
  const targets = [];
  let offset = 0;
  let total = null;
  let runtime = null;
  do {
    const result = await runJsonCandidateChain(docxCandidates, [
      'find-template-migration-targets',
      sourcePath,
      baselinePath,
      sourceChoiceId,
      branch,
      text ?? '-',
      String(offset),
      '100',
    ]);
    const page = migrationTargetPage.parse(result.json);
    if (page.sourceChoiceId !== (sourceChoiceId === '-' ? null : sourceChoiceId) || page.branch !== branch) {
      throw new Error('Migration target page identity does not match the requested current source and action');
    }
    if (total !== null && page.total !== total) {
      throw new Error('Migration target page total changed while reading current alternatives');
    }
    if (page.offset !== offset || page.targets.length === 0 && offset < page.total) {
      throw new Error('Migration target pagination did not advance');
    }
    runtime ??= commandRuntime(result);
    total = page.total;
    targets.push(...page.targets);
    offset += page.targets.length;
  } while (offset < total);
  return { runtime, targets };
}

function migrationChoiceRelevance(source, target) {
  const sourceContext = source.context ?? {};
  const targetContext = target.context ?? {};
  const sourceRows = sourceContext.sameRowTexts ?? [];
  const targetRows = targetContext.sameRowTexts ?? [];
  const sourceHeaders = sourceContext.tableHeaderTexts ?? [];
  const targetHeaders = targetContext.tableHeaderTexts ?? [];
  const mainText = textSimilarity(source.text, target.text);
  const tableIdentity = Math.max(
    textSimilarity(sourceContext.columnHeaderText, targetContext.columnHeaderText),
    maximumPairSimilarity(sourceHeaders, targetHeaders));
  const neighborhood = Math.max(
    textSimilarity(sourceContext.previousText, targetContext.previousText),
    textSimilarity(sourceContext.nextText, targetContext.nextText),
    maximumPairSimilarity(sourceRows, targetRows),
    maximumPairSimilarity([source.text], targetRows),
    maximumPairSimilarity(sourceRows, [target.text]));
  return source.kind === 'table-cell' && target.kind === 'table-cell'
    ? [tableIdentity, mainText, neighborhood]
    : [mainText, neighborhood, tableIdentity];
}

function compareRelevance(left, right) {
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) return left[index] - right[index];
  }
  return 0;
}

function maximumPairSimilarity(leftValues, rightValues) {
  let maximum = 0;
  for (const left of leftValues ?? []) {
    for (const right of rightValues ?? []) {
      maximum = Math.max(maximum, textSimilarity(left, right));
    }
  }
  return maximum;
}

function textSimilarity(leftValue, rightValue) {
  const left = normalizeSearchText(leftValue);
  const right = normalizeSearchText(rightValue);
  if (!left || !right) return 0;
  if (left === right) return 1;
  if (Math.min(left.length, right.length) >= 3 && (left.includes(right) || right.includes(left))) {
    return Math.min(left.length, right.length) / Math.max(left.length, right.length);
  }
  const leftPairs = characterPairs(left);
  const rightPairs = characterPairs(right);
  if (leftPairs.size === 0 || rightPairs.size === 0) return 0;
  let shared = 0;
  for (const pair of leftPairs) if (rightPairs.has(pair)) shared += 1;
  return 2 * shared / (leftPairs.size + rightPairs.size);
}

function normalizeSearchText(value) {
  return typeof value === 'string'
    ? value.normalize('NFKC').toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '')
    : '';
}

function characterPairs(value) {
  const characters = [...value];
  if (characters.length < 2) return new Set(value ? [value] : []);
  const pairs = new Set();
  for (let index = 0; index + 1 < characters.length; index += 1) {
    pairs.add(characters[index] + characters[index + 1]);
  }
  return pairs;
}

function migrationTargetBranch(action, sourceKind) {
  if (action === 'place-content') return sourceKind === 'media' ? 'copy-media' : 'copy-text';
  if (action === 'keep-template-content') return 'retain-target';
  if (action === 'keep-template-label') return 'retain-target-label';
  if (action === 'select-template-option') return 'choice-selection';
  throw new Error(`Unsupported migration target action: ${action}`);
}

function catalogReferenceToken(catalog) {
  return createHash('sha256')
    .update(`${catalog.sourceSha256}\n${catalog.baselineSha256}`)
    .digest('hex')
    .slice(0, 8);
}

function choiceRef(catalog, items, id, prefix) {
  const index = items.findIndex(item => item.id === id);
  if (index < 0) throw new Error(`Migration choice is not bound to the current catalog: ${id}`);
  return `${prefix}${index + 1}-${catalogReferenceToken(catalog)}`;
}

function choiceIdFromRef(catalog, items, ref, prefix, label) {
  if (!ref.startsWith(prefix)) {
    throw invalidMigrationInput(`invalid migration ${label} ref: ${ref}`);
  }
  const [ordinal, token] = ref.slice(1).split('-');
  if (token !== catalogReferenceToken(catalog)) {
    throw invalidMigrationInput(`stale migration ${label} ref: ${ref}`);
  }
  const item = items[Number(ordinal) - 1];
  if (!item) throw invalidMigrationInput(`unknown migration ${label} ref: ${ref}`);
  return item.id;
}

const migrationActionCodes = new Map([
  ['place-content', 'P'],
  ['keep-template-label', 'L'],
  ['select-template-option', 'O'],
]);
const migrationActionsByCode = new Map([...migrationActionCodes].map(([action, code]) => [code, action]));

function migrationAlternativeRef(catalog, source, action, target) {
  const sourceOrdinal = catalog.sources.findIndex(item => item.id === source.id) + 1;
  const targetOrdinal = catalog.targets.findIndex(item => item.id === target.id) + 1;
  const code = migrationActionCodes.get(action);
  if (sourceOrdinal < 1 || targetOrdinal < 1 || !code) {
    throw new Error('Migration alternative is not bound to the current catalog');
  }
  return `S${sourceOrdinal}-${code}${targetOrdinal}-${migrationAlternativeToken(catalog, source, action, target)}`;
}

function migrationAlternativeToken(catalog, source, action, target) {
  return createHash('sha256')
    .update(`${catalog.sourceSha256}\n${catalog.baselineSha256}\n${source.id}\n${action}\n${target.id}`)
    .digest('hex')
    .slice(0, 8);
}

function migrationChoiceFromAlternativeRef(catalog, ref) {
  const match = /^S([1-9][0-9]*)-([PLO])([1-9][0-9]*)-([0-9a-f]{8})$/.exec(ref);
  if (!match) throw invalidMigrationInput(`invalid migration alternative ref: ${ref}`);
  const source = catalog.sources[Number(match[1]) - 1];
  const target = catalog.targets[Number(match[3]) - 1];
  const action = migrationActionsByCode.get(match[2]);
  if (!source || !target || !action) {
    throw invalidMigrationInput(`unknown migration alternative ref: ${ref}`);
  }
  if (match[4] !== migrationAlternativeToken(catalog, source, action, target)) {
    throw invalidMigrationInput(`stale or invalid migration alternative ref: ${ref}`);
  }
  return { sourceChoiceId: source.id, action, targetChoiceId: target.id };
}

function publicMigrationChoice(catalog, choice, prefix) {
  const { id, ...visible } = choice;
  const items = prefix === 'S' ? catalog.sources : catalog.targets;
  return {
    ...visible,
    allowedActions: visible.allowedActions.filter(action => action !== 'keep-template-content'),
    ref: choiceRef(catalog, items, id, prefix),
  };
}

function migrationQueryResult({ runtime, catalog, view, action, source, items, offset, total, alternatives = false }) {
  return {
    tool: 'docx_query_migration_choices',
    runtime,
    sourceSha256: catalog.sourceSha256,
    baselineSha256: catalog.baselineSha256,
    view,
    action,
    source: source ? publicMigrationChoice(catalog, source, 'S') : null,
    items: alternatives
      ? items
      : items.map(item => publicMigrationChoice(catalog, item, view === 'sources' ? 'S' : 'T')),
    page: {
      offset,
      returned: items.length,
      total,
      hasMore: offset + items.length < total,
    },
  };
}

async function docxMigrateTemplate(args) {
  return runTemplateMigrationCommand('docx_migrate_template', 'migrate-template', args);
}

async function docxVerifyMigration(args) {
  return runTemplateMigrationCommand('docx_verify_migration', 'verify-template-migration', args);
}

function invalidMigrationInput(message) {
  return Object.assign(new Error(message), { code: -32602 });
}

function completeMigrationChoices(catalog, choices) {
  const sources = new Map(catalog.sources.map(source => [source.id, source]));
  const seen = new Set();
  const completed = choices.map(rawChoice => {
    const choice = rawChoice.alternativeRef
      ? migrationChoiceFromAlternativeRef(catalog, rawChoice.alternativeRef)
      : {
          sourceChoiceId: choiceIdFromRef(catalog, catalog.sources, rawChoice.sourceRef, 'S', 'source'),
          action: rawChoice.action,
        };
    const { sourceChoiceId } = choice;
    const source = sources.get(sourceChoiceId);
    if (!source) throw invalidMigrationInput(`unknown migration source id: ${sourceChoiceId}`);
    if (seen.has(sourceChoiceId)) throw invalidMigrationInput(`duplicate migration source id: ${sourceChoiceId}`);
    seen.add(sourceChoiceId);
    if (!source.allowedActions.includes(choice.action)) {
      throw invalidMigrationInput(`migration action ${choice.action} is not allowed for source id: ${sourceChoiceId}`);
    }
    return source.requiredCardinality === 'all'
      ? { ...choice, cardinality: 'all' }
      : choice;
  });
  const missing = [...sources.keys()].filter(sourceChoiceId => !seen.has(sourceChoiceId));
  if (missing.length > 0) {
    throw invalidMigrationInput(`migration choices must cover every source id; missing ${missing.length}`);
  }
  return completed;
}

function completeTemplateCleanup(catalog, cleanup, choices) {
  const targets = new Map(catalog.targets.map(target => [target.id, target]));
  const claimedTargets = new Set(choices.flatMap(choice => choice.targetChoiceId ? [choice.targetChoiceId] : []));
  const seen = new Set();
  return cleanup.map(rawCleanup => {
    const targetChoiceId = choiceIdFromRef(catalog, catalog.targets, rawCleanup.targetRef, 'T', 'target');
    if (!seen.add(targetChoiceId)) throw invalidMigrationInput(`duplicate migration cleanup target id: ${targetChoiceId}`);
    if (!targets.get(targetChoiceId)?.allowedActions.includes('template-cleanup')) {
      throw invalidMigrationInput(`migration cleanup is not allowed for target id: ${targetChoiceId}`);
    }
    return { targetChoiceId, scope: rawCleanup.scope };
  }).filter(cleanupChoice => !claimedTargets.has(cleanupChoice.targetChoiceId));
}

async function runTemplateMigrationCommand(tool, command, args) {
  const source = requireString(args.source, 'source');
  const baseline = requireString(args.baseline, 'baseline');
  const output = requireString(args.output, 'output');
  if (!Array.isArray(args.choices)) {
    throw Object.assign(new Error('choices must be an array'), { code: -32602 });
  }
  const catalogResult = await runJsonCandidateChain(docxCandidates, ['list-template-migration-choices', source, baseline]);
  const catalog = migrationCatalog.parse(catalogResult.json);
  const choices = completeMigrationChoices(catalog, args.choices);
  const templateCleanup = Array.isArray(args.templateCleanup)
    ? completeTemplateCleanup(catalog, args.templateCleanup, choices)
    : [];
  const payload = {
    schema: 'tiwater.docx.template-migration-business-choices/v1',
    choices,
    ...(templateCleanup.length > 0 ? { templateCleanup } : {}),
  };
  return withTempJsonFile(payload, async choicesPath => {
    const result = await runJsonCandidateChain(
      docxCandidates,
      [command, source, baseline, choicesPath, output],
      { allowedExitCodes: [0, 1] });
    if (result.json === null) {
      const detail = result.stderr.trim() || result.stdout.trim() || 'no diagnostic output';
      throw new Error(`${result.command} ${command} returned no JSON receipt (exit ${result.code}): ${detail}`);
    }
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
