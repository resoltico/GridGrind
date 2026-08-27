---
afad: "5.0.1"
version: "0.74.0"
domain: DEPENDENCY_AUTOMATION_POLICY
updated: "2026-08-28"
route:
  keywords: [gridgrind, dependabot, dependency, dependency-update, approval, gate, release-hygiene]
  questions: ["how are GridGrind dependency updates approved", "what should happen to open Dependabot PRs before a release", "can GridGrind auto-merge Dependabot PRs"]
---

# Dependency Automation Policy

GridGrind requires an explicit human decision for every Dependabot update. Automation may prepare a pull request, but it never authorizes a merge.

## Triage

| Tier | Trigger | Deadline | Required action |
|:-----|:--------|:---------|:----------------|
| Security | Security advisory on a direct or transitive dependency | Within seven calendar days of opening | Review, verify, then merge or reject with a documented reason. |
| Regular | Non-security routine update | Before the next release | Review during release hygiene; merge, close, or explicitly keep open. |
| Major | Semver-major update in any ecosystem | Before the next release | Treat as a compatibility change; verify affected public and build surfaces explicitly. |

## Approval And Closeout

- Never merge a Dependabot PR without a successful CI `Gate` check.
- A Docker base-image update also requires a successful `Docker smoke` check.
- A Gradle update that changes Apache POI, Log4j, or Jackson also requires the `Check` job that exercises CLI-contract verification and Jazzer regression.
- A GitHub Actions update requires the pinned workflow SHA to match the adopted release tag; verify it with `gh api repos/<owner>/<repo>/git/ref/tags/<tag>`.
- Before ending a release session, every open Dependabot PR must be merged, closed with an explicit reason and branch deletion, or deliberately kept open with a still-valid reason recorded on the PR.
- Never retag or amend a published release to absorb a dependency update. Post-publication dependency work lands through a new `main` pull request.
