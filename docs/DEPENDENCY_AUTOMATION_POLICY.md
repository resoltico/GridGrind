# Dependency Automation Policy

GridGrind verifies dependency freshness directly with the publisher or vendor. Dependabot is an
advisory signal, not an inventory or release authority. Update authorization and merge
authorization are separate; automation never authorizes a merge.

## Inventory and verification

- Inventory every catalog and resolved locked coordinate, settings/build plugin, wrapper,
  workflow action, image/frontend/helper, contributor tool and OS snapshot.
- Check the publisher's stable releases and matching artifacts directly. Exclude prereleases
  unless explicitly adopted. Record retrieval time, exact identity/checksum and source URL.
- Keep related families aligned. Treat Gradle's embedded libraries and compiler/tool-private
  dependencies as their owner's supported tuple; document concrete owner/compatibility reasons
  for retaining older private components and exercise the actual owning path.
  [Private dependency tuples](DEPENDENCY_TUPLES.md) documents the current owner boundaries.
- Regenerate locks through resolution, verify a clean locked restore, reconcile new/removed
  graph nodes, and inspect packaged runtime bytes and legal/source inventory independently.
- Use exact immutable workflow/image/tool pins and authenticated fixed OS snapshots. A pinned
  installer that internally chooses `latest` does not make the installed tools reproducible.

## Triage

| Tier | Trigger | Deadline | Required action |
|:-----|:--------|:---------|:----------------|
| Security | Security advisory on a direct or transitive dependency | Within seven calendar days of opening | Review, verify, then merge or reject with a documented reason. |
| Regular | Non-security routine update | Before the next release | Review during release hygiene; merge, close, or explicitly keep open. |
| Major | Semver-major update in any ecosystem | Before the next release | Treat as a compatibility change; verify affected public and build surfaces explicitly. |

## Approval And Closeout

- Never merge a Dependabot PR without a successful CI `Gate` check.
- A Docker base-image update also requires a successful `Docker smoke` check.
- A Gradle update that changes Apache POI, Log4j, or Jackson also requires the `Check` job that exercises deterministic tests, the golden corpus and fuzz seed replay.
- A GitHub Actions update requires the pinned workflow SHA to match the adopted release tag; inspect it with `gh api repos/<owner>/<repo>/git/ref/tags/<tag>` and dereference annotated tags to the underlying commit before comparing.
- Before ending a release session, every open Dependabot PR must be merged, closed with an explicit reason and branch deletion, or deliberately kept open with a still-valid reason recorded on the PR.
- Never retag or amend a published release to absorb a dependency update. Post-publication dependency work lands through a new `main` pull request.
