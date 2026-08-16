# Template migration decision lifecycle contract review

| Item | Contract |
| --- | --- |
| Owner | Published `tiwater-docx` template-migration decision lifecycle. |
| Primary input | One current source DOCX, one selected baseline DOCX, and one provider-owned run-local decision draft bound to both document hashes. |
| Machine output | Progress for the next unresolved source, available targets for that source and decision branch, atomic decision or revision receipts, and one final validated migration plan or closed local-review result. |
| Invariants | Target lookup and decision recording use the same branch language. Lookup excludes targets already claimed by the draft, except that distinct typed facts may share one retained-label paragraph or table cell; duplicate claims on one typed field fail closed. A decision is validated against freshly reopened documents before it is committed; rejection leaves the draft unchanged and the same source current. Revising one recorded source is atomic and never requires replaying other decisions. Resolution accepts only a complete draft whose recorded decisions have already passed the same semantic resolver. Keeping a template label preserves its target-owned presentation and fields while deterministically migrating uniquely delimited identifier, version, and date facts from the selected current parent; missing or ambiguous target slots fail closed. |
| Consumer | An Agent reads the current source and bounded target observations, chooses business meaning, records or revises one decision, and continues from returned progress. Deterministic provider code owns identities, draft mutation, semantic admission, closure, and plan generation. |
| Compatibility | Existing disposition-first target lookup and explicit-source recording remain accepted. New callers use the aligned branch syntax and the purpose-named revision command. |
| Non-goals | Choosing business meaning, ranking targets, reading Lucid scenario knowledge, generating business values, authoring operations, editing documents, or deciding delivery. |
