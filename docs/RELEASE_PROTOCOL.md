---
afad: "3.5"
version: "0.49.0"
domain: RELEASE_PROTOCOL
updated: "2026-04-21"
route:
  keywords: [gridgrind, release, gh, github-cli, java26, gradlew, tag, ci, container, docker]
  questions: ["how do I release gridgrind", "what is the gridgrind release procedure", "how do I verify java before a gridgrind release", "how do I publish a gridgrind tag release"]
---

# Release Protocol

The entire release flow is driven by the GitHub CLI (`gh`). Every step that touches GitHub —
PRs, CI status, merges, releases, workflow monitoring — uses `gh`, not the GitHub web UI.

**BEFORE DOING ANYTHING ELSE**, run both checks:

```bash
gh --version
gh auth status
```

If either command fails — `gh` is not installed, or `gh auth status` reports "not logged in" —
**STOP IMMEDIATELY**. Do not attempt any further steps. Notify the user:

> GitHub CLI (`gh`) is not available or not authenticated. The release procedure cannot
> continue. Please install `gh` and run `gh auth login` (this requires browser interaction
> and possibly 2FA, so it must be done by you, not by me). Once `gh auth status` reports a
> logged-in account with repo access, tell me to resume.

Do not attempt to resolve missing `gh` or authentication failures autonomously.

---

### Step 1 — Pre-flight: verify release readiness

Before running any build or release command, verify the local Java and Gradle runtime:

```bash
command -v java
java --version
./gradlew --version --console=plain
```

Requirements before continuing:

- `command -v java` must not be `/usr/bin/java`.
- `java --version` must report Java 26.
- Use `./gradlew`, never Brew `gradle`, for every repo build or release step.

Then run `./check.sh`. It must exit 0. If it fails, fix all failures before proceeding.

Then verify every item in this checklist. All must be true before any commit or tag:

- `gradle.properties` `version=` equals the target release version exactly (e.g. `0.3.0`).
- `CHANGELOG.md` has a `## [X.Y.Z] - YYYY-MM-DD` section (not `[Unreleased]`) with at least
  one entry.
- `CHANGELOG.md` link footer has:
  - `[Unreleased]: .../compare/vX.Y.Z...HEAD`
  - `[X.Y.Z]: .../compare/vPREV...vX.Y.Z`
- All Markdown docs that carry AFAD frontmatter — `README.md`, `PATENTS.md`, `jazzer/README.md`,
  and every `docs/*.md` file — have `version:` set to the target version.
- `README.md` does not reference any prior version's container tags.
- All example JSON files use the current wire names and field shapes for this version.
- GitHub repository settings are still aligned with this procedure:
  - default branch is `main`
  - `delete_branch_on_merge` is enabled
  - `main` is protected with admin enforcement
  - required status checks are exactly `Check` and `Docker smoke`

Before cutting the release branch, enumerate open PRs so dependency-automation work is never
surprise-discovered after publication:

```bash
gh pr list --state open \
  --json number,title,url,headRefName,mergeStateStatus,isDraft,author,statusCheckRollup
```

If any open PR is authored by `dependabot[bot]`, decide up front whether it changes release
machinery or release-critical dependencies. If it does, land or reject it before cutting the
release branch. If it does not, carry that decision forward and complete Step 10 before ending
the release session.

### Step 2 — Commit on a release branch

`main` is branch-protected. Never attempt `git push origin main` directly — it will be
rejected and wastes time. Always commit on a release branch.

```bash
git checkout -b release/X.Y.Z
git add <every modified file that belongs in the release — never .codex/>
git status --short
git diff --cached --name-status
git diff --cached --stat
git commit -m "release: bump version to X.Y.Z"
git push origin release/X.Y.Z
```

Treat staging as a handoff checkpoint, not a formality. Before committing:

- `git status --short` must show no intended release file left unstaged or untracked.
- `git diff --cached --name-status` must show the exact file set expected for the release.
- `git diff --cached --stat` must confirm that the staged payload includes both versioning or
  docs changes and every production, test, workflow, or script change that belongs in the
  release.

If the staged diff is incomplete or includes unintended files, fix the branch before committing.
Do not rely on memory alone to decide what is in the release.

### Step 3 — Open PR and wait for CI

```bash
gh pr create \
  --title "release: bump version to X.Y.Z" \
  --base main \
  --head release/X.Y.Z \
  --body "..."
```

Note the PR number returned. Then poll CI until the required check passes:

```bash
gh pr diff <N> --name-only
gh pr view <N> --json number,state,mergeStateStatus,statusCheckRollup,url
```

Treat the PR itself as a second scope-verification checkpoint:

- `gh pr diff <N> --name-only` must match the intended release file set.
- If the PR diff is missing files or includes unintended files, fix the release branch before
  waiting on CI or merging.
- Every new commit pushed to the release branch reopens both the Step 2 staging checkpoint and
  this PR diff checkpoint. Re-verify both after each fix commit.

Do not proceed until **every** required job in workflow `CI` has `"conclusion": "SUCCESS"`.
At the time of writing that means both `Check` and `Docker smoke`. If any required job fails,
fix the failure, push to the release branch, and wait again — do not merge a red PR.

### Step 4 — Merge PR, wait for main CI, and verify the merge handoff

```bash
gh pr merge <N> --merge --admin --delete-branch --subject "release: bump version to X.Y.Z (#N)"
git checkout main
git pull
./scripts/verify-release-merge-handoff.sh
gh pr view <N> --json number,state,mergedAt,headRefName,baseRefName,url
```

The `--admin` flag uses administrator privileges to bypass branch-protection requirements,
specifically the review-approval rule that GitHub prevents the PR author from satisfying.
This is the GitHub-intended escape hatch for single-owner repositories where an agent drives
the release end-to-end. CI status checks (`Check` and `Docker smoke`) remain the authoritative
quality gate; the review requirement adds no signal in a solo-owner workflow.

Requirements before continuing:

- PR state is `MERGED`.
- `mergedAt` is populated.
- Local `main` contains the merge commit you expect.
- Local `main` exactly matches `origin/main`.
- The merged `main` commit already has successful `Check` and `Docker smoke` runs from workflow
  `CI`. `./scripts/verify-release-merge-handoff.sh` is the authoritative gate for this handoff.
- The remote release branch is deleted by the merge step.

GitHub auto-delete on merge should also be enabled at the repository level. `--delete-branch`
remains mandatory here so the release handoff stays self-contained even if the repo setting is
misconfigured or temporarily changed.

Do not create the tag until this merge-handoff verifier passes. Tagging before the merged
`main` commit's own `CI` run finishes will cause the Release and Container workflows to fail
their publication-safety gate on the first attempt.

The release branch must not be left behind. If the local `release/X.Y.Z` branch still exists
after the merge, delete it manually with:

```bash
git branch -d release/X.Y.Z
```

### Step 5 — Create the tag, push it, and verify the tag handoff

```bash
git tag vX.Y.Z
git push origin vX.Y.Z

REPO=$(gh repo view --json nameWithOwner -q .nameWithOwner)
gh api "repos/$REPO/git/ref/tags/vX.Y.Z"
```

Do not proceed until the remote tag ref exists. Never infer a successful tag push from the
absence of a local git error alone — verify the remote ref through GitHub.

The tag push is what triggers the Release and Container workflows. The PR merge alone does
not. These are two separate actions — both are required.

Step 4's merge-handoff verifier is a hard prerequisite here. Only tag the commit that is both
checked out locally and already green on `origin/main`.

If either publication workflow later needs a targeted rerun against the existing tag, use:

```bash
gh workflow run release.yml -f release_tag=vX.Y.Z
gh workflow run container.yml -f release_tag=vX.Y.Z
```

Never create a second tag or move an existing release tag just to retry CI.

Both publication workflows re-run the release-candidate gate before any build or publish step.
An existing-tag rerun is expected to fail unless all of the following are still true:

- `gradle.properties` in the checked-out tag still reports `version=X.Y.Z`
- the workflow checkout matches the exact remote `vX.Y.Z` tag commit
- that tag commit remains reachable from the default branch (`main`)
- that exact commit already has successful `Check` and `Docker smoke` runs from workflow `CI`

### Step 6 — Branch hygiene

After the merge and tag push, clean up stale remote-tracking refs and verify that no historical
release branches remain on GitHub.

```bash
REPO=$(gh repo view --json nameWithOwner -q .nameWithOwner)
git remote prune origin
gh api "repos/$REPO/branches" --paginate --jq '.[].name'
```

Requirements:

- No `release/X.Y.Z` branch may remain on GitHub after the merge.
- No historical `release/` branches may remain on GitHub. If any are present, delete them:

```bash
git push origin --delete release/A.B.C
```

- No fully merged local `release/` branches may remain. Delete them:

```bash
git branch -d release/A.B.C
```

Do not leave release-branch leftovers behind locally or remotely. Branch hygiene is part of the
release procedure, not optional cleanup.

Open maintenance branches such as Dependabot are handled separately in Step 10. Do not treat a
non-`release/` branch as automatically acceptable just because Step 6 only hard-fails
`release/*` leftovers.

### Step 7 — Monitor workflows with duplicate-run awareness

```bash
gh run list --workflow=release.yml --branch=vX.Y.Z --event=push --limit=10
gh run list --workflow=container.yml --branch=vX.Y.Z --event=push --limit=10
```

Do not assume there is exactly one run per workflow. A single tag push may produce multiple runs
for the same workflow. Treat the workflow boundary as a **handoff checkpoint**:

1. Enumerate **all** runs for `release.yml` on `vX.Y.Z`.
2. Enumerate **all** runs for `container.yml` on `vX.Y.Z`.
3. Inspect each run that is not `completed/success` with:

```bash
gh run view <run-id> --log-failed
```

4. Verify the external GitHub state directly before deciding the release is failed.

Rules:

- Never treat one failed run as authoritative if another sibling run for the same tag succeeded.
- Never re-run blindly. First inspect whether the desired state already exists.
- A release-workflow failure with `Release.tag_name already exists` is **not** automatically a
  release failure. It may mean a sibling run already created the release successfully.
- Only classify the release workflow as failed if **no** run produced the required external
  state and direct GitHub inspection confirms that state is absent or incomplete.

Fix the root cause only after the direct-state inspection proves the release or container state
is actually missing or incorrect. Coordinate with the user if the failure is in CI infrastructure
outside this codebase.

When duplicate runs are observed for the same workflow, tag, and commit, classify the **source**
of the duplicate dispatch separately from the **safety** of the publication system:

- The source may be a user- or tool-driven duplicate tag push, a client retry, or a GitHub
  Actions delivery anomaly.
- Unless GitHub audit evidence proves which one occurred, treat the source as externally
  ambiguous. Do not present guesswork as certainty.
- Inside this repository, the required engineering response is still deterministic: the workflows
  must remain safe under duplicate dispatch. Concurrency, idempotent publication, and direct
  post-publication verification are mandatory.

### Step 8 — Verify the GitHub Release handoff

Do not infer release publication from workflow success alone. Verify the release object directly:

```bash
gh release view vX.Y.Z --json tagName,isDraft,isPrerelease,publishedAt,url,assets
```

Requirements:

- The release exists for tag `vX.Y.Z`.
- `isDraft` is `false`.
- `isPrerelease` is `false` unless the target release is intentionally a prerelease.
- The fat JAR asset is present (currently `gridgrind.jar`).

If these conditions are satisfied, the GitHub Release handoff is complete even if an additional
duplicate release workflow run failed after the release was already created.

The release workflow is expected to perform this same verification internally after publication.
Before publication, that workflow also black-box verifies the packaged fat JAR with
`./scripts/verify-cli-contract.sh jar ./cli/build/libs/gridgrind.jar` so the shipped `--help`
plus `--print-protocol-catalog`, `--print-task-catalog`, `--print-task-plan`, `--print-goal-plan`,
and `--doctor-request` surfaces cannot drift from the core contract silently. The operator-side
`gh release view` check remains mandatory because workflow success is still not the authoritative
state.

### Step 9 — Verify public availability

Do not declare the release done until the GitHub Release exists and the following anonymous
Docker verification succeeds. Use the repository verifier so you are testing the public surface,
not cached owner credentials, while still targeting the active local Docker engine selected in
the current shell:

```bash
gh release view vX.Y.Z                                                # GitHub Release with fat JAR
./scripts/verify-container-publication.sh ghcr.io/resoltico/gridgrind X.Y.Z
```

The verifier script must confirm both the exact `X.Y.Z` tag and `latest` are anonymously pullable,
that both `docker run ... --version` results match the two-line product header for the target
release version exactly — `GridGrind X.Y.Z` on the first line and the product description on the
second — and that both published tags still expose the expected `--help`,
`--print-protocol-catalog`, `--print-task-catalog`, `--print-task-plan`, `--print-goal-plan`, and
`--doctor-request` contract. A successful `docker pull` alone is not sufficient verification. In
particular: a multi-arch `docker pull` can succeed even when the platform
manifests have been deleted — the index manifest is still present but the image is not actually
runnable. The `docker run --version` plus CLI-contract checks remain the definitive test.

If the verifier script fails, inspect the published state, fix the release surface, and rerun the
same verification command. Do not switch to the operator's normal Docker config as a fallback.

These checks are a second handoff checkpoint. Workflow success is not enough; public pull and run
behavior is the authoritative state.

The container workflow now also publishes OCI provenance and SBOM attestations for the multi-arch
image. Treat those attestations as part of the release artifact set, but do not let them replace
the required anonymous pull-and-run verification above.

The container registry retains the last 5 releases. Only `X.Y.Z` and `latest` tags are
published per release; there is no `X.Y` floating tag.

Only after the release view check and the verifier script both succeed report to the user: the
release is publicly available.

The container workflow is expected to perform the exact-tag and `latest` pull-and-run verification
internally after publication. The operator-side verification remains mandatory because public
availability, not workflow success, is the authoritative state.

### Step 10 — Triage Dependabot PRs and clear dependency-automation leftovers

After the public release is verified, do not end the release session while open Dependabot PRs are
still sitting untriaged. Release hygiene includes dependency-automation hygiene.

Re-enumerate all open PRs and identify Dependabot-owned entries directly from GitHub metadata:

```bash
gh pr list --state open \
  --json number,title,url,headRefName,mergeStateStatus,isDraft,author,statusCheckRollup
```

Treat any PR whose `author.login` is `dependabot[bot]` as in scope for this step, even if it was
already reviewed during Step 1. Step 1 creates the release-time decision; Step 10 closes the loop
before the release session is allowed to end.

For each open Dependabot PR, inspect the exact payload and its current gate status:

```bash
gh pr diff <N> --name-only
gh pr view <N> --json number,title,state,mergeStateStatus,statusCheckRollup,url
```

Rules:

- If the PR is wanted, mergeable, and already green on the required `CI` checks, merge it
  immediately and delete its branch:

```bash
gh pr merge <N> --merge --admin --delete-branch --subject "<title> (#<N>)"
```

- If the PR is stale, superseded by `main`, intentionally rejected, or replaced by a different
  change path, close it explicitly and delete its branch:

```bash
gh pr close <N> --comment "Superseded or intentionally rejected during release hygiene." --delete-branch
```

- If the PR needs follow-up work before it is acceptable, do that work as a normal post-release
  change on `main` and then land or replace the Dependabot PR. Do not leave a green but
  unattended Dependabot PR parked indefinitely just because the release itself already shipped.

- Never retag, amend, or move the just-published release tag to absorb a Dependabot change. The
  published release remains immutable. Dependabot resolution is post-release `main` hygiene.

- There is no "ignore it and leave the branch there" option. Every open Dependabot PR must end
  this step in exactly one of these states:
  - merged and branch deleted
  - closed and branch deleted
  - consciously kept open with an explicit still-valid reason

After each merge or close, resync and re-check GitHub branch state:

```bash
git checkout main
git pull
git remote prune origin
REPO=$(gh repo view --json nameWithOwner -q .nameWithOwner)
gh api "repos/$REPO/branches" --paginate --jq '.[].name'
```

Requirements before declaring the release session complete:

- No stale Dependabot PR may remain open without an explicit keep-open decision.
- No merged or closed Dependabot branch may remain on GitHub.
- Any remaining non-`main` branch on GitHub must correspond to an intentional still-open PR that
  was reviewed during this step and deliberately kept alive.
