# Office MCP boundary

## Goal

The published Office MCP server is the Agent-facing surface for Word, Excel,
and PowerPoint technical capabilities. It is installed from a package registry
and uses the exact published Office runtimes. Production does not require this
repository or discover Office work through command-line help.

The server uses the official MCP SDK for protocol negotiation, stdio framing,
argument validation, and structured-result validation. The adapter declares
each public tool once; it does not maintain a second protocol implementation or
hand-written validator for the same contract.

Large observations are durable artifacts, not model messages. The caller
chooses a new run-local JSON path; the adapter writes the complete provider
result once and returns only its path, hash, and size. A choice catalog remains
an opaque durable artifact. A bounded semantic query reads the same current
source and baseline directly, so the Agent never parses or relays catalog
storage.

Command-line programs remain implementation and compatibility surfaces. They
may expose diagnostic commands, but those commands do not become an Agent
workflow.

## Template migration

Template migration has four public operations:

- list the current migration choices;
- query a bounded source page or the targets for one source choice;
- migrate the current document after receiving one typed batch of choices;
- independently verify the result from the same current inputs and choices.

The first operation writes every source item that still needs business judgment
and the current baseline targets to an evidence artifact. The query operation
re-observes the same current source and baseline, then returns a bounded source
page, provider-compatible targets for one opaque source identity and business
action, or cleanup targets. Literal visible-text filtering is optional.
Compatible targets remain pageable and complete; their display order uses only
current visible text and local document context such as neighboring text and
table headers. That search order is not a business recommendation or selection.
The Agent selects identities and actions without parsing the artifact format or
reconstructing technical target compatibility. It does not author document
content, selectors, coordinates, edit operations, plans, or intermediate files.

The third operation validates the complete batch, derives the technical plan,
edits a temporary copy of the baseline, reads the result back, and returns the
output and a complete receipt. A missing, stale, duplicate, or incompatible
choice fails before mutation. Genuine business ambiguity remains local review;
it does not erase determinate choices or become a guessed mapping.

The fourth operation independently repeats admission and readback from the
current source, baseline, choices, and output. It does not trust the migration
operation's embedded verdict.

All three operations publish MCP input and output schemas and return structured
content. Text content is a human-readable projection of the same result, not a
second contract.

The Agent-facing server does not expose target search, incremental draft,
record, revise, replay, plan construction, apply, legacy candidate commands, or
low-level mutation payloads such as style-id maps. Those may remain command-line
compatibility or diagnostic operations, but the Agent does not assemble them.

## Decision boundary

Deterministic code owns document inventory, exact technical matches, identity
and uniqueness checks, target occupancy, operation generation, mutation, and
readback. The Agent owns only a business choice among alternatives presented by
the tool and allowed by the current scenario knowledge.

Forty unresolved items are forty choices, not a reason to invent another
abstraction. This boundary deliberately does not introduce region expansion, a
knowledge compiler, a scenario schema registry, or an incremental decision
protocol. If a future reusable Office behavior cannot be expressed by the
existing choices, that behavior receives a separate capability review; a new
scenario alone does not change the Office interface.

Scenario knowledge remains natural-language business authority. Its quality can
change which business choice the Agent makes, but malformed prose does not alter
the tool protocol or cause the Office runtime to guess fields, commands, or JSON
shapes.

## Ownership

- The scenario package owns business identity, allowed ambiguity, and terminal
  meaning.
- The Agent chooses among the tool's typed current alternatives.
- The published Word runtime owns document observation, including local table
  context, deterministic planning, editing, and readback.
- The Office MCP adapter publishes the typed surface, orders complete technical
  alternatives for progressive discovery, and invokes the exact installed
  runtime; it does not interpret scenario rules or choose a target.
- Lucid owns workflow ordering, evidence handoff, delivery verification, and
  platform boundaries.

## Rejected alternatives

Incremental find-record-revise was rejected because it exposes bookkeeping,
multiplies tool calls with document size, and asks the Agent to manage provider
state. The bounded query is stateless discovery: it records no decision and the
final migration still receives one complete choice batch. Region-level
automatic expansion was rejected because it requires a new semantic inference
model and risks moving scenario meaning into the Word runtime. A single opaque
migration operation with no independent verifier was rejected because it
weakens evidence and makes producer errors self-validating.

## Compatibility and acceptance

- Existing command-line callers continue to work.
- A clean machine can install and start the MCP server without a repository
  checkout or source fallback.
- Tool discovery lists only supported Agent-facing operations.
- Rejected choices leave source, baseline, output, and prior evidence unchanged.
- Different valid document shapes use the same public operations.
- Adding a scenario does not add an Office tool.
