# Text MCP boundary

Status: stable capability contract

The published Text MCP process observes plain-text bytes and decoded line
coordinates. It is distributed with the existing server package and reuses its
transport and immutable-artifact mechanisms, but it is a separate process and
tool surface from Office document capabilities.

The provider accepts only explicitly supported plain-text extensions and only
encodings that can be decoded losslessly: UTF-8 with or without its BOM, and
UTF-16 little- or big-endian with the corresponding BOM. Invalid byte
sequences, unsupported encodings, binary control content, and ambiguous
extensionless or binary formats fail without producing an observation.

An inspection binds the exact source bytes, reports decoding facts and line
count, returns a bounded opening identity, and retains the complete inspection
as an immutable artifact. A line read selects one explicit zero-based offset
and bounded limit. Every returned line identity belongs to the exact source
hash, and every bounded page reports its remaining count and continuation.

The provider does not parse key-value pairs, sections, records, markup,
Markdown structure, field meaning, or business rules. It does not repair,
normalize, rewrite, or transcode source content.
