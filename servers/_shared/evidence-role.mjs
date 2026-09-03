export const evidenceRoleMetadataKey = 'x-tiwater-evidence-role';
export const evidenceRoleSchema = 'tiwater.provider-evidence-role/v1';

const evidenceRoles = new Set([
  'document-observation',
  'final-readback',
  'native-render',
]);

export function evidenceRoleMetadata(role) {
  if (!evidenceRoles.has(role)) throw new Error(`unsupported-evidence-role:${role}`);
  return {
    [evidenceRoleMetadataKey]: {
      schema: evidenceRoleSchema,
      role,
    },
  };
}

export function assertEvidenceToolContract(tool, expectedRole) {
  const metadata = tool?._meta?.[evidenceRoleMetadataKey];
  if (!metadata || Object.keys(metadata).sort().join(',') !== 'role,schema'
      || metadata.schema !== evidenceRoleSchema || metadata.role !== expectedRole) {
    throw new Error(`evidence-role-metadata-invalid:${tool?.name || 'unnamed'}`);
  }
  const bindings = fileBindings(tool.inputSchema);
  if (!bindings.some(binding => binding.role === 'read')) {
    throw new Error(`evidence-role-source-binding-missing:${tool.name}`);
  }
  if (!tool.outputSchema || typeof tool.outputSchema !== 'object') {
    throw new Error(`evidence-role-output-schema-missing:${tool.name}`);
  }

  if (expectedRole === 'native-render') {
    assertAnnotations(tool, false, false);
    if (!bindings.some(binding => binding.role === 'write' && binding.effect)
        || !bindings.some(binding => binding.role === 'write' && !binding.effect)
        || !requiredArtifact(tool.outputSchema, 'receipt')) {
      throw new Error(`native-render-evidence-contract-invalid:${tool.name}`);
    }
    return;
  }

  assertAnnotations(tool, true, true);
  const writes = bindings.filter(binding => binding.role === 'write');
  if (writes.length < 1 || writes.some(binding => binding.effect)) {
    throw new Error(`read-evidence-artifact-binding-invalid:${tool.name}`);
  }
  if (!sourceIdentity(tool.outputSchema) || !requiredArtifact(tool.outputSchema, 'artifact')) {
    throw new Error(`read-evidence-output-binding-invalid:${tool.name}`);
  }
  if (expectedRole === 'document-observation') {
    const identity = requiredObject(tool.outputSchema, 'identity') || requiredObject(tool.outputSchema, 'summary');
    if (!identity || unboundedArrays(identity).length > 0) {
      throw new Error(`document-observation-identity-unbounded:${tool.name}`);
    }
  } else if (!requiredObject(tool.outputSchema, 'receipt')) {
    throw new Error(`final-readback-receipt-missing:${tool.name}`);
  }
}

function assertAnnotations(tool, readOnlyHint, idempotentHint) {
  const annotations = tool.annotations || {};
  if (annotations.readOnlyHint !== readOnlyHint
      || annotations.idempotentHint !== idempotentHint
      || annotations.destructiveHint !== false
      || annotations.openWorldHint !== false) {
    throw new Error(`evidence-role-annotations-invalid:${tool.name}`);
  }
}

function fileBindings(schema) {
  const bindings = [];
  function visit(node) {
    if (!node || typeof node !== 'object' || Array.isArray(node)) return;
    if (node['x-tiwater-file-role'] === 'read' || node['x-tiwater-file-role'] === 'write') {
      bindings.push({
        role: node['x-tiwater-file-role'],
        effect: node['x-tiwater-file-role'] === 'write' && node['x-tiwater-file-effect'] !== false,
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

function requiredObject(schema, property) {
  const value = schema?.properties?.[property];
  return schema?.required?.includes(property) && value?.type === 'object' ? value : null;
}

function artifactShape(schema) {
  const candidate = schema?.anyOf?.find(value => value?.type === 'object') || schema;
  return candidate?.type === 'object'
    && ['path', 'sha256', 'bytes'].every(property => candidate.required?.includes(property));
}

function requiredArtifact(schema, property) {
  return schema?.required?.includes(property) && artifactShape(schema?.properties?.[property]);
}

function sourceIdentity(schema) {
  if (requiredArtifact(schema, 'source')) return true;
  const sources = schema?.properties?.sources;
  return schema?.required?.includes('sources') && sources?.type === 'array'
    && Number.isInteger(sources.maxItems) && artifactShape(sources.items);
}

function unboundedArrays(schema, location = '$', found = []) {
  if (Array.isArray(schema)) {
    schema.forEach((child, index) => unboundedArrays(child, `${location}[${index}]`, found));
    return found;
  }
  if (!schema || typeof schema !== 'object') return found;
  if (schema.type === 'array' && !Number.isInteger(schema.maxItems)) found.push(location);
  for (const [key, child] of Object.entries(schema)) unboundedArrays(child, `${location}.${key}`, found);
  return found;
}
