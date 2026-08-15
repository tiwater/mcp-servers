# Template migration object selector contract review

| Item | Contract |
| --- | --- |
| Owner | Published `tiwater-docx` template-migration observation capability. |
| Primary input | One current source DOCX and one selected baseline DOCX; both are inventoried independently. |
| Machine output | Each observable migration object may carry a `Selector` that identifies that object uniquely within its own current inventory using the existing semantic-selector fields. Candidate discovery returns one `RequiredDecisions` entry for every source not already covered by the automatic exact mapping, optional mechanically observed target suggestions, and unclaimed baseline observations; it never returns an executable plan. |
| Invariants | An observation contains only kind, scope, visible text, and an optional unique selector. It contains no object id, index, coordinate, scenario value, or inferred business mapping. Resolving a present selector against the same inventory yields exactly that object. The unclaimed baseline set is recomputed from the same hash-bound plan and omits every already claimed baseline object. Reordering non-semantic package metadata does not change it. |
| Consumer | An Agent considers every required source, chooses source and baseline objects by business meaning, and copies published selectors into an existing semantic candidate. Zero or multiple suggestions do not create a review terminal. The Agent does not join technical object ids back to the full analysis inventory. The resolver independently reopens both current documents and resolves those selectors. |
| Compatibility | The legacy anchor-gap and exact-plan commands remain callable but are hidden from new discovery. Existing analysis, resolver, and candidate branches remain valid. |
| Non-goals | Choosing a source-to-target mapping, deciding a disposition, deriving business values, accepting object ids or coordinates in candidates, editing documents, or replacing independent resolution and readback. |
