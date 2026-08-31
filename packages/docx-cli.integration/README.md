# DOCX table operation matrix

Run the reusable table conformance matrix from the repository root:

```sh
npm run test:docx-table-operations
```

The matrix uses scenario-independent documents and exercises the published table
operations across flat, horizontal-merge, vertical-merge, mixed-merge,
rectangular-merge, irregular-grid, multi-paragraph, and nested tables. It checks
reads after every mutation, Word identity uniqueness, table-grid and merge-owner
invariants, OpenXML validity, legal operation sequences, boundary insertion and
deletion, and atomic failure behavior.

When a scenario exposes a table-operation defect, first reduce it to a generic
table shape and operation sequence in this matrix. The regression must reproduce
the failure without scenario names, work-item identifiers, customer values, or
expected scenario answers. Fix the owning DOCX implementation only after the
generic regression fails, then keep the case permanently in this matrix.
