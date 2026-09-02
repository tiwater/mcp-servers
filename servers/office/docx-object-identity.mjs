export function compactDocxObjectIdentity(object) {
  return {
    address: object.address,
    parentAddress: object.parentAddress ?? null,
    kind: object.kind,
    textPreview: object.textPreview ?? null,
    gridSpan: object.gridSpan ?? null,
    verticalMerge: object.verticalMerge ?? null,
    verticalTextAlignment: object.verticalTextAlignment ?? null,
  };
}
