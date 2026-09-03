import { evidenceRoleMetadataKey } from './evidence-role.mjs';

export const effectKindMetadataKey = 'x-tiwater-effect-kind';
export const effectKindSchema = 'tiwater.provider-effect-kind/v1';
export const documentRevisionRoleKey = 'x-tiwater-document-revision-role';

const effectKinds = new Set([
  'document-mutation',
  'source-conversion',
  'native-render',
]);

export function effectKindMetadata(kind) {
  if (!effectKinds.has(kind)) throw new Error(`unsupported-effect-kind:${kind}`);
  return {
    [effectKindMetadataKey]: {
      schema: effectKindSchema,
      kind,
    },
  };
}

export function assertEffectKindToolContract(tool, expectedKind = undefined) {
  const metadata = tool?._meta?.[effectKindMetadataKey];
  const bindings = fileBindings(tool?.inputSchema);
  const effectiveWrites = bindings.filter(binding => binding.role === 'write' && binding.effect);
  const readOnly = tool?.annotations?.readOnlyHint === true;
  const currentDocumentBindings = bindings.filter(binding => binding.revisionRole === 'current');
  const invalidRevisionBindings = bindings.filter(binding => binding.revisionRole !== undefined
    && (binding.revisionRole !== 'current' || binding.role !== 'read'));
  if (invalidRevisionBindings.length > 0) {
    throw new Error(`document-revision-role-invalid:${tool?.name || 'unnamed'}`);
  }

  if (readOnly || effectiveWrites.length === 0) {
    if (metadata !== undefined) throw new Error(`effect-kind-unexpected:${tool?.name || 'unnamed'}`);
    if (currentDocumentBindings.length !== 0) {
      throw new Error(`document-revision-role-unexpected:${tool?.name || 'unnamed'}`);
    }
    return null;
  }
  if (!metadata || Object.keys(metadata).sort().join(',') !== 'kind,schema'
      || metadata.schema !== effectKindSchema || !effectKinds.has(metadata.kind)) {
    throw new Error(`effect-kind-metadata-invalid:${tool?.name || 'unnamed'}`);
  }
  if (expectedKind !== undefined && metadata.kind !== expectedKind) {
    throw new Error(`effect-kind-mismatch:${tool?.name || 'unnamed'}:${metadata.kind}`);
  }
  if (!bindings.some(binding => binding.role === 'read')) {
    throw new Error(`effect-kind-source-binding-missing:${tool?.name || 'unnamed'}`);
  }
  if (metadata.kind === 'document-mutation') {
    if (currentDocumentBindings.length !== 1) {
      throw new Error(`document-mutation-current-binding-invalid:${tool?.name || 'unnamed'}`);
    }
    if (effectiveWrites.length !== 1) {
      throw new Error(`document-mutation-output-binding-invalid:${tool?.name || 'unnamed'}`);
    }
  } else if (currentDocumentBindings.length !== 0) {
    throw new Error(`document-revision-role-unexpected:${tool?.name || 'unnamed'}`);
  }

  const evidenceRole = tool?._meta?.[evidenceRoleMetadataKey]?.role;
  if ((metadata.kind === 'native-render') !== (evidenceRole === 'native-render')) {
    throw new Error(`native-render-effect-evidence-mismatch:${tool?.name || 'unnamed'}`);
  }
  if (metadata.kind !== 'native-render' && evidenceRole !== undefined) {
    throw new Error(`effect-kind-evidence-role-conflict:${tool?.name || 'unnamed'}`);
  }
  return metadata.kind;
}

function fileBindings(schema) {
  const bindings = [];
  function visit(node) {
    if (!node || typeof node !== 'object' || Array.isArray(node)) return;
    if (node['x-tiwater-file-role'] === 'read' || node['x-tiwater-file-role'] === 'write') {
      bindings.push({
        role: node['x-tiwater-file-role'],
        effect: node['x-tiwater-file-role'] === 'write'
          && node['x-tiwater-file-effect'] !== false,
        revisionRole: node[documentRevisionRoleKey],
      });
    }
    for (const child of Object.values(node.properties || {})) visit(child);
    if (node.items) visit(node.items);
    for (const keyword of ['allOf', 'anyOf', 'oneOf']) {
      for (const child of node[keyword] || []) visit(child);
    }
  }
  visit(schema);
  return bindings;
}
