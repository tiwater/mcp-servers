# Template migration object selector contract review

| Item | Contract |
| --- | --- |
| Owner | Published `tiwater-docx` template-migration observation capability. |
| Primary input | One current source DOCX and one selected baseline DOCX; both are inventoried independently. |
| Machine output | Each observable migration object may carry a `Selector` that identifies that object uniquely within its own current inventory using the existing semantic-selector fields. Exact-text and anchor-gap plans repeat the compact current observations needed for each unresolved source, any mechanically discovered baseline option, every unclaimed non-empty baseline paragraph or table cell, and selectable child runs of those cells. |
| Invariants | An observation contains only kind, scope, visible text, and an optional unique selector. It contains no object id, index, coordinate, scenario value, or inferred business mapping. Resolving a present selector against the same inventory yields exactly that object. The unclaimed baseline set is recomputed from the same hash-bound plan and omits every already claimed baseline object. Reordering non-semantic package metadata does not change it. |
| Consumer | An Agent reads the unresolved plan itself, chooses source and baseline objects by business meaning, and copies published selectors into an existing semantic candidate. It does not join technical object ids back to the full analysis inventory. The resolver independently reopens both current documents and resolves those selectors. |
| Compatibility | All added observations are optional additive fields. Existing analysis and plan consumers and all existing candidate branches remain valid. |
| Non-goals | Choosing a source-to-target mapping, deciding a disposition, deriving business values, accepting object ids or coordinates in candidates, editing documents, or replacing independent resolution and readback. |
