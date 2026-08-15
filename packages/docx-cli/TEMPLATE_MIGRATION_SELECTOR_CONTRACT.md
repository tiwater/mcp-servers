# Template migration object selector contract review

| Item | Contract |
| --- | --- |
| Owner | Published `tiwater-docx` template-migration observation capability. |
| Primary input | One current DOCX opened by `analyze-template-migration`; source and baseline are inventoried independently. |
| Machine output | Each observable migration object may carry a `Selector` that identifies that object uniquely within its own current inventory using the existing semantic-selector fields. |
| Invariants | The selector contains no object id, index, coordinate, scenario value, or inferred business mapping. Resolving it against the same inventory yields exactly that object. Reordering non-semantic package metadata does not change it. If no supported semantic selector is unique, `Selector` is absent. |
| Consumer | An Agent may choose source and baseline objects by business meaning and copy their published selectors into an existing semantic candidate. The resolver independently reopens both current documents and resolves those selectors. |
| Compatibility | `Selector` is an optional additive observation field. Existing analysis consumers and all existing candidate branches remain valid. |
| Non-goals | Choosing a source-to-target mapping, deciding a disposition, deriving business values, accepting object ids or coordinates in candidates, editing documents, or replacing independent resolution and readback. |

