# Office MCP boundary

Status: frozen target design

## Purpose

The published Office MCP server is the Agent-facing technical surface for
Word, Excel, PowerPoint, and native Office rendering. Its public concepts are
the document objects already defined by Open XML: parts, paragraphs, runs,
tables, rows, cells, drawings, worksheets, ranges, slides, and shapes.

The server does not define a document migration language, a business workflow,
or a second document object model. New business packages compose the same
Office capabilities and do not add an Office tool or field.

Every capability has one published machine input, one machine output, one
owner, and one stated non-goal. The MCP adapter publishes that contract and
invokes the exact installed runtime without translating it into another
request shape. Command-line entry points may expose the same contract for
diagnostics; they are not a second Agent-facing protocol.

## Authority and decision flow

- Open XML defines document structure and physical identity.
- The provider observes that structure and performs technical operations on it.
- The business-rule owner defines meaning, source and target roles, and
  acceptance semantics.
- The Agent reads current documents, applies those rules, and chooses
  the source, target, value, and terminal disposition.
- The orchestrator sequences calls and preserves accepted provider calls and
  evidence unchanged.
- Luna independently reviews final business content and rendered appearance.

Information flows in that direction only. The provider never proposes a
business mapping, and the orchestrator never reinterprets an Office object or
rewrites a provider call.

## Document revision and object identity

Every observation is bound to the exact input bytes and published runtime
version. Every returned object identity belongs to that revision.

Word identities address an Open XML story part and an object within that part,
including the main document, headers, footers, comments, text boxes, footnotes,
and endnotes.
Workbook identities use the workbook revision, worksheet identity, and native
A1 cell or range address. Presentation identities use the presentation
revision, slide part, and native shape identity. Child objects retain their
native parent relationship.

Provider-issued identities are opaque to callers but remain traceable to these
native coordinates in evidence. They are not semantic selectors and do not
contain business roles. Mutation creates a new revision; callers re-observe it
instead of reusing stale identities. The provider rejects stale, missing,
ambiguous, cross-document, or wrong-kind identities before writing output.

## Observation capabilities

Observation is progressively disclosed through four orthogonal capabilities:

- inspect the package and return its revision, parts, top-level structure, and
  bounded summaries;
- list objects of a requested native kind and scope;
- find literal current text within a native scope;
- read the selected current objects in full technical detail.

Inspection may also write a complete durable artifact for evidence. Paging and
result limits bound transport only; they do not silently discard matches or
turn excess results into a business decision. A bounded result reports the
remaining count and a continuation mechanism.

Observation reports facts present in the current document. It does not infer
field names, headings, semantic roles, source-target compatibility, preferred
targets, cleanup candidates, or review status.

## Mutation capabilities

Mutation consists of fixed Open XML actions. Each public operation has one verb,
optionally batches independent objects of the same valid kind, and produces a
new document plus a receipt. There is no generic operation discriminator,
expression language, plan compiler, or provider-owned workflow.

The public operation list is closed over these verbs:

| Operation | Valid native objects and effect |
| --- | --- |
| set content | Replace content in selected Word text containers, workbook cells, or presentation text and table cells without replacing their containers. |
| copy content | Copy current content from selected source containers to selected target containers; target structure and formatting remain authoritative unless a fixed formatting mode is explicitly selected. |
| insert object | Insert selected existing paragraphs, tables, rows, cells, runs, or text into a selected native container with their required parts and relationships. |
| move object | Move selected objects within one current document without changing their content. |
| delete object | Delete selected objects and remove relationships that become unreferenced. |
| set properties | Set published typed Open XML properties for selected text, paragraph, table, row, cell, section, worksheet, slide, or shape objects. |
| merge cells | Merge a selected rectangular Word-table or worksheet-cell range when the native format permits it. |
| split cells | Split a selected merged Word-table or worksheet-cell range when the native format permits it. |
| replace media | Replace selected drawing or picture bytes while preserving the selected container and its declared geometry. |
| apply layout | Apply a caller-selected Word style, worksheet presentation policy, or presentation master/layout without selecting it on business grounds. For presentations, a selected master/layout/theme may come from another current revision; the provider imports its complete related-part dependency closure while preserving target slide content, count, and order. |
| refresh fields | Refresh caller-selected Word field scopes through the declared native Writer backend and preserve unrelated document content. |
| validate package | Validate package integrity and the exact requested technical postconditions. |
| convert format | Convert a supported Office-family input through its declared native application. |
| render document | Render a supported Office-family input through its declared native application with provenance. |

Each format publishes only the operations valid for its native objects. The
format prefix and action verb identify the public operation; object kind is
validated from the provider identity rather than selected through a nested
operation payload. Adding support for another typed property or native object
does not create a new verb.

A direct copy receives current source identities and current target identities.
The Agent supplies their relationship; the provider performs the package work
needed to preserve runs, styles, table grids, merges, formulas, drawings, and
relationships according to the selected fixed action. This keeps Open XML
mechanics out of Agent glue without moving semantic selection into code.

Table transformation is a composition of the same native-object actions, not a
separate table language. Copying row or grid-column objects establishes target
structure; setting or copying cell content fills that structure; setting typed
properties changes presentation; merging or splitting cells changes topology;
moving or deleting objects changes order and extent. None of these actions
infers a row group, column meaning, source-target mapping, or business value.

Cross-document copy and presentation layout import bind both current revisions
explicitly. They are valid only for the fixed operations that declare source
and target documents; an identity cannot otherwise be spliced into another
document call. Presentation layout import copies system parts and dependencies,
not template slide content.

The provider may reject a technically impossible action with a precise reason.
It must not guess another target, alter the requested business value, expand a
selection semantically, or convert rejection into a review decision.

## Execution and receipts

Mutation preflights every selected object against the same current revision,
applies the fixed action to a temporary output, validates the package, and only
then publishes the output. Failure leaves the input and requested output
unchanged.

The receipt binds the provider version, input revision, accepted call, output
revision, applied objects, and technical postconditions. The orchestrator stores
the receipt and exact accepted call; it does not derive a second operation
representation. Large read results and receipts may be durable artifacts, but their
paths are transport details rather than document identities.

Provider validation proves package integrity and requested technical effects.
It does not prove that the Agent selected the correct business source, target,
value, or disposition. The independent reviewer determines those from current
authoritative inputs, business rules, final readback, and native render.

## Native rendering

Rendering accepts one current Word, Excel, or PowerPoint-family document and
uses the matching declared WPS application. Its receipt binds the input,
backend, output PDF, page count, byte count, and content hashes. Rendering does
not inspect appearance or make a delivery decision.

Legacy `.xls` conversion belongs to the workbook conversion capability and
uses the declared ET/WPS spreadsheet backend before Open XML workbook
operations begin. OCR remains a PDF capability and is not performed by an
Office document tool.

## Non-goals

Office capabilities do not own:

- business fields, identities, mappings, acceptance rules, or current-job
  answers;
- inferred candidates, recommended alternatives, cleanup suggestions, source
  conservation verdicts, or human-review terminals;
- migration choices, drafts, plans, compilers, registries, replay protocols, or
  independent business validators;
- orchestration lifecycle, evidence closure, platform state, or delivery status;
- hidden fallback parsing, prompt repair, or business-package compatibility.

## Enforced isolation

The boundary is a release condition rather than a documentation convention.

- Provider processes receive document bytes or paths, provider-issued revision
  and object identities, fixed technical values, and artifact destinations.
  Business-rule packages, orchestration workspaces, and delivery state are not
  mounted, passed, imported, or fetched.
- Office MCP registrations are generated from the provider-owned machine
  contracts. The adapter cannot hand-author another tool, request schema,
  description, or compatibility branch.
- Public requests and results contain only native object identities, typed
  technical values, artifacts, revisions, receipts, and technical failures.
  Business choices and workflow terminals are not valid contract types.
- The release gate checks the package dependency graph, generated MCP manifest,
  public schemas, and a clean installed-package launch with no business-rule or
  orchestration repository available. A dependency, contract, or tool outside
  this boundary stops publication.
- A new public verb requires a separate design-only boundary change and closure
  review before implementation. A new object kind or typed Open XML property
  under an existing verb does not create another workflow or protocol.

Source-word checks may detect accidental business terminology, but they are
only a backstop. Isolation, closed machine types, generated registration, and
the release gate are the enforcing mechanisms.

## Compatibility and closure

Existing fixed Open XML actions may be retained when they fit the boundary and
adopt the single revision identity. Business decision, candidate, plan,
builder, apply-workflow, and business-validator surfaces do not fit it and are
removed rather than preserved as an alternate Agent path.

An interface change is justified only by a reusable Open XML technical
responsibility that cannot be composed from these families. A business package,
issue, current job, customer value, filename, table coordinate, or known answer can
never justify one. Different valid document shapes must use the same published
capabilities, and adding business packages must not increase the capability families.
