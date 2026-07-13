# Runtime evidence trust contracts

These versioned contracts are the shared trust boundary between published
`tiwater-*` runtimes and workflow orchestrators. They define evidence shapes;
they do not identify Office or PDF formats, choose business meaning, or mutate
documents.

## Capability descriptor

Every published runtime exposes a zero-input `capabilities --json` command whose
output conforms to `runtime-capabilities.schema.json`. The descriptor names the
package, executable runtime, evidence schema, supported file kinds, command
surface, and one non-mutating identify probe. The identify probe has exactly
three outcomes:

- `supported`: the runtime recognized one of its declared kinds from runtime-owned
  signature evidence;
- `unsupported`: the source was read successfully but did not match a declared
  kind;
- `failed`: the probe could not produce a trustworthy supported/unsupported
  decision.

`unsupported` is not an error and must never be collapsed into `failed`. Failed
results retain a source identity only when those exact bytes were successfully
hashed; the failure must name a later probe stage and must not claim a matched
kind or signature. A
capability descriptor declares only its own probes and kinds. It does not declare
the orchestrator's required probe set.

## Evidence envelope

`runtime-evidence-envelope.schema.json` binds one probe result to the exact
source bytes with path, byte length, SHA-256, and content id. Package, runtime,
and evidence-schema identities are required independently so consumers can reject
version drift.

The `file` member contains runtime-owned file kind, media type, and signature
evidence. A supported result requires a matched signature. Unsupported and failed
results retain the same source identity but do not claim a file kind.

`objects[]` exposes runtime evidence nodes and containment. A runtime may use a
native id only when the format actually supplies one. Nodes without a native id
must use an explicitly derived identity with deterministic inputs. A derived
identity must never populate `nativeId`. Every non-root object names its parent.
JSON Schema enforces the local shape; consumers must additionally enforce unique
object ids, resolvable parent references, and acyclic containment.

The envelope's `artifact` hashes canonical JSON bytes of `payload`, not the
envelope containing the hash. This avoids a circular self-hash. The caller that
retains the complete envelope separately records the retained envelope path,
size, and SHA-256 in its run registry.

Canonical JSON v1 recursively sorts unique object names by ordinal UTF-16 code
units, preserves array order, writes UTF-8 without insignificant whitespace,
uses lowercase hexadecimal digits in `\\u` escapes, rejects duplicate keys, and accepts
only null, booleans, strings, arrays, objects, and cross-language safe integer
numbers from `-9007199254740991` through `9007199254740991`. Exact decimal
business values must remain strings. Mathematically integral JSON forms such as
`1.0` and `1e0` normalize to `1`. This deliberately narrow
number rule keeps the C#, Python, and JavaScript artifact bytes identical without
letting language-specific floating-point rendering change an artifact id.

## Edit report

`edit-report.schema.json` describes a complete runtime mutation attestation.
There is exactly one `operations[]` entry for each request entry, in request
order, with a zero-based index and matching type. Each entry retains the complete
requested payload and the complete payload actually applied, plus targets,
warnings, and errors. Rejected or failed entries use `appliedPayload: null` and
must explain the failure. Applied and no-op entries retain a non-null applied
payload; no-op does not mean the payload may be omitted.

`summary` is derived only from operation statuses. The report must be checked
against a separately retained authoritative request; reconstructing that request
from `operations[].requestedPayload` would make the report its own oracle.
Schema validation cannot prove
ordering, count derivation, digest recomputation, parent resolution, or equality
between the request and report. Independent validators must recompute those
relationships, as demonstrated by `runtime-contracts.test.mjs`.

## Versioning and non-goals

All three schemas begin at `1.0.0`. Additive optional fields may use a compatible
minor schema version only when old consumers remain correct. Removing fields,
changing meanings, weakening required provenance, or changing canonical bytes
requires a new incompatible schema version.

These contracts intentionally do not:

- parse or sniff Office/PDF bytes in shared code;
- invent native ids for tables, rows, cells, runs, or detected PDF structures;
- allow model-authored values or coordinates to become executable payloads;
- make an edit report its own correctness oracle;
- replace caller-side hashing of retained descriptors, envelopes, requests, or
  output artifacts.

The JSON files under `fixtures/` are synthetic contract examples only. They are
not format fixtures and contain no scenario or customer data.
