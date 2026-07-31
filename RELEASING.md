# Release process

Published `tiwater-*` packages must come from the repository's `main` branch.
The merge commit on `main` is the release source authority.

## Required sequence

1. Make the implementation and package-version change on a feature branch.
2. Run the package's direct tests and any affected integration tests from a clean worktree.
3. Push the feature branch and open a pull request to `main`.
4. Review the complete diff and test evidence. Do not approve a release from uncommitted files.
5. Merge the pull request into `main`.
6. The `Publish Shared Runtimes` workflow runs from that `main` merge commit and publishes the new package version.
7. Wait for the workflow to succeed and for its registry-installability checks to pass.
8. Independently install the exact published version into a clean tool path before updating a consumer repository's version pin.

## Prohibited release paths

- Do not publish from a feature branch, pull-request ref, dirty worktree, local package directory, or manually modified deployment checkout.
- Do not manually upload NuGet or PyPI artifacts.
- Do not use `workflow_dispatch` on a non-`main` ref.
- Do not update a consumer's pinned version before the `main` workflow and clean installation proof both succeed.
- Do not create a second version bump merely to recover from registry indexing delay.

`workflow_dispatch` is an infrastructure-retry mechanism only. It must select
`main` and retry the exact committed version already present there; it is not a
substitute for merging a release pull request.

## Release evidence

Record these items in the pull request or delivery handoff:

- merged pull request and exact `main` merge SHA;
- package identifier and version;
- direct and affected integration-test results;
- publish workflow URL and terminal result;
- registry artifact hash when available;
- clean-path installation result for the exact version.

If the `main` workflow does not start or fails, stop and repair the workflow or
release commit. Do not publish the same source from another ref.
