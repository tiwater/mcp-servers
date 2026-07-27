# Test fixtures

`lucid.provider-contract-manifest.schema.json` is a byte-for-byte copy of the
Lucid schema-set 15 provider contract manifest schema, copied from the
lucid-docs repository:

- Source: `plugins/lucid/workflow/schema-sets/15/provider-contract-manifest.schema.json`
  (lucid-docs working tree, copied 2026-07-27)
- SHA-256 of the copied bytes:
  `7eddd13c38eb9b61d82787292b3b46433caafc29ea31f04ffae14725d60e14bc`

The docx/xlsx/pptx manifest tests validate produced
`lucid.provider-contract-manifest` documents against these exact bytes with a
real Draft 2020-12 JSON Schema evaluator (JsonSchema.Net). The fixture is not
packed into any nupkg; it lives outside `contracts/` on purpose.
