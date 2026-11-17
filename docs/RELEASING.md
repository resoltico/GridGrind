# Releasing

[Release](../.github/workflows/release.yml) is a dispatch-only workflow on protected `main`.
`prepare` builds and checks an untagged source commit; `publish-prepared` publishes an explicitly
reviewed immutable artifact pair. Tags do not trigger a build or publication. Release controls,
source and the protected main head must be the same commit. All releases share one serialized
workflow concurrency group.

## Prepare the source and candidate

Commit the intended product, documentation, dependency and legal inputs before preparation.
Move the reviewed Unreleased bullets into one exact `## [X.Y.Z] - YYYY-MM-DD` changelog section,
leaving an empty Unreleased section above it. Release notes come from that section. Documentation-only
changes do not need release-note bullets. Local builds use `0.0.0-dev`; preparation supplies
`-Pversion=X.Y.Z` to the declared build, archive and verification tasks.

From the matching clean checkout, set the intended numeric stable version and dispatch:

```sh
repo=resoltico/GridGrind
version=X.Y.Z # Replace with the intended unpublished numeric version.
source_commit=$(git rev-parse HEAD)
test "$source_commit" = "$(gh api "repos/$repo/branches/main" --jq .commit.sha)"
gh workflow run release.yml --repo "$repo" --ref main \
  -f operation=prepare -f version="$version" -f source_commit="$source_commit"
```

The workflow refuses a dirty checkout, another source/workflow commit, unprotected or moved
main, an existing version tag, or an existing release/draft. It checks identity before expensive
work. Its preparation and QA jobs have read-only publication permissions.

`Prepare candidate` runs release-control tests, the deterministic check/build graph, a bounded
active fuzz run and a bounded real mutation control. `prepareReleaseNotes` validates and writes
the dated changelog section in the same deterministic Gradle graph before assembly. It derives
the resolved component inventory
with `:cli:writeRuntimeComponentInventory`: `cli/build/runtime-component-inventory.json` contains
portable component/legal facts; `cli/build/runtime-component-files.json` maps build-local files
and is not a retained public inventory. Exact versions and checksums remain in the
[authoritative pins](DEVELOPER.md#toolchain-and-dependencies).

Preparation builds retained AMD64 and ARM64 OCI runtime graphs with provenance and SPDX SBOMs,
loads those same roots, and checks the JAR, launcher archives, capsule, stock provider classes,
legal bytes and source correspondence. It exercises the runtime container suite for each selected
platform and the contributor container. The explicit borrowed-image form is:

    ./gradlew containerTest -PcontainerImage=<retained-local-image> -PcontainerPlatform=linux/arm64

`containerPlatform` also accepts `linux/amd64` and requires a supplied image. This form verifies
the selected JVM architecture without rebuilding or deleting the borrowed image.

Only completed checks permit sealing and uploading `delivery-X.Y.Z-RUN-ATTEMPT`. The separate
`Bind candidate QA receipt` job downloads that exact service artifact ID, checks the actual ZIP
and payload, then uploads `artifact-qa-X.Y.Z-RUN-ATTEMPT`. Both use immutable service artifacts
with seven-day retention. They are CI candidates visible to authenticated artifact readers;
preparation creates no public tag, release or registry image. Failure diagnostics are separate
from the two publication roles.

## Download and independently accept the pair

Wait for the entire preparation run to succeed. Record its run ID, current attempt, and both
artifact IDs and service digests from GitHub, not merely artifact names:

    gh run list --repo "$repo" --workflow release.yml --branch main --commit "$source_commit"
    gh api "repos/$repo/actions/runs/$candidate_run_id"
    gh api --paginate "repos/$repo/actions/runs/$candidate_run_id/artifacts"

Set `candidate_run_id`, `candidate_run_attempt`, `delivery_id`, `delivery_digest`, `qa_id` and
`qa_digest` to the reviewed service values. Digests use `sha256:` followed by 64 lowercase hex
digits. With a credential in `GH_TOKEN` that can read Actions artifacts, use the trusted checkout's
helper and a destination that does not exist:

```sh
python3 -B scripts/release_pair.py download \
  --repo "$repo" --version "$version" --source "$source_commit" \
  --run-id "$candidate_run_id" --attempt "$candidate_run_attempt" \
  --delivery-id "$delivery_id" --delivery-digest "$delivery_digest" \
  --qa-id "$qa_id" --qa-digest "$qa_digest" --output reviewed-release
```

[The pair validator](../scripts/release_pair.py) verifies the workflow, branch, source, successful
jobs, current attempt, exact role IDs/digests, service ZIP bytes, candidate inventory and QA
binding. `reviewed-release/delivery/` holds `candidate.json`, release assets, OCI archives and
check evidence. `reviewed-release/qa/artifact-qa.json` binds that candidate to the delivery
service artifact. `reviewed-release/reviewed-pair.json` binds both roles and their inner hashes.
Never execute scripts supplied by downloaded payloads; use controls from the reviewed checkout.

Independently review the downloaded contents, applicable changed contracts, platform behavior,
actual test discovery and required failure controls. Use **that downloaded**
`delivery/assets/gridgrind.jar` for any additional runtime acceptance. Record what was exercised,
its outcomes and remaining limits in retained review evidence. The automated artifact receipt
binds earlier CI checks; it does not replace independent acceptance.

After the applicable review passes, the reviewer creates `independent-acceptance.json` using the
exact [acknowledgement contract](../scripts/release_candidate.py):

| Field | Meaning |
| --- | --- |
| `formatVersion`, `role` | `1` and `independent-acceptance` |
| `pairSha256`, `jarSha256` | The reviewed canonical pair and downloaded JAR |
| `status` | `PASS`, supported by the independent review evidence |
| `evidenceSha256` | SHA-256 of the retained independent QA evidence |

The acknowledgement is bounded to 32 KiB. Its shape, canonical digest and pair/JAR binding are
checked automatically; those checks cannot establish that the declared independent work happened.
The reviewer owns the evidence behind PASS. This is an explicit dispatcher acknowledgement,
not a manual GitHub environment gate or an automatic publisher trigger.

## Publish exactly the accepted delivery

Generate dispatch inputs from the verified pair and the completed acknowledgement. Run this
read-only preparation in the trusted checkout; it validates the acknowledgement before writing
`publish-inputs.json`:

```sh
python3 - <<'PY'
import json
from pathlib import Path
import sys
sys.path.insert(0, 'scripts')
from release_candidate import canonical, independent_acceptance, read_json, sha_bytes, sha_file
record = read_json(Path('reviewed-release/reviewed-pair.json'))
pair = record['pair']
identity = pair['identity']
ack = read_json(Path('independent-acceptance.json'))
ack_sha = sha_bytes(canonical(ack))
jar_sha = sha_file(Path('reviewed-release/delivery/assets/gridgrind.jar'))
independent_acceptance(ack, ack_sha, record['pairSha256'], jar_sha)
inputs = {
    'operation': 'publish-prepared', 'version': identity['version'],
    'source_commit': identity['sourceCommit'],
    'candidate_run_id': str(identity['runId']),
    'candidate_run_attempt': str(identity['runAttempt']),
    'delivery_artifact_id': str(pair['delivery']['id']),
    'delivery_artifact_digest': pair['delivery']['digest'],
    'candidate_manifest_sha256': pair['candidateSha256'],
    'qa_artifact_id': str(pair['qa']['id']), 'qa_artifact_digest': pair['qa']['digest'],
    'qa_receipt_sha256': pair['receiptSha256'], 'reviewed_pair_sha256': record['pairSha256'],
    'independent_acceptance': canonical(ack).decode('utf-8'),
    'independent_acceptance_sha256': ack_sha,
}
Path('publish-inputs.json').write_text(json.dumps(inputs, indent=2) + '\n', encoding='utf-8')
PY
gh workflow run release.yml --repo "$repo" --ref main --json < publish-inputs.json
```

The [publisher](../scripts/publish_prepared.py) downloads and verifies the pair again. It checks
current protected-main controls, all committed input hashes, owned source bytes, OCI graphs,
SBOMs and the independent acknowledgement before any tag, release, registry or retention mutation.
It does not rebuild the product, regenerate sources, rerun the preparation suites or execute
artifact-supplied code. The installed OCI transfer client is checksum-pinned in
[transfer-tools.properties](../release/transfer-tools.properties).

Publication creates or verifies the version tag, stages the exact assets, then exposes a non-latest
public prerelease with source assets. Anonymous source retrieval must match the reviewed bytes.
The data-only source image is published and checked before the versioned runtime graph. After
runtime/source verification and conservative retention, the publisher advances `latest` and
promotes the GitHub release to stable/latest. There is no atomic transaction across GitHub and
GHCR. A failure can leave an inspectable prerelease or partial registry state; it does not establish
a completed stable release. Publication receipts and diagnostics are retained separately.

## Expiry, changes and retries

Missing, deleted or expired artifacts refuse publication. The seven-day candidate retention is
not an indefinite recovery store. A new preparation attempt needs both fresh artifact roles and
fresh independent acceptance; roles or acknowledgements cannot be mixed across attempts. A
partial failed-job rerun cannot substitute an older delivery into a newer QA attempt. Use a full
preparation attempt when replacing an unaccepted candidate.

Moved main, changed controls or changed product inputs require a complete new preparation and
review. If publication already created a tag or release/draft, inspect that state: preparation
refuses to treat it as a new untagged version. Do not retag or overwrite public assets to bypass
that refusal. A publication retry is a new `publish-prepared` dispatch naming the same still-valid
reviewed pair. Exact matching draft/prerelease state can resume; mismatched assets or registry
roots refuse. An already matching stable release receives read-only verification rather than
replacement. Inspect ambiguous network failures before retrying; publication mutations are not
blindly repeated.

## Container source delivery

[Runtime pins](../release/container-runtime.properties) and the [Dockerfile](../Dockerfile) own one
immutable Temurin JDK/Ubuntu 26.04 family for build, contributor and application images. Fixed
Ubuntu snapshot indexes remain authenticated; only historical metadata expiry checks are disabled.

Every delivery includes `gridgrind-X.Y.Z-sources.tar.gz`, created from the exact committed Git tree
before a tag exists and checked against every tracked blob and executable mode. It is separate
from `gridgrind-container-sources-X.Y.Z.tar.gz` and its manifest/checksums, which bind matching
Temurin source/build inputs and installed Ubuntu binary/source versions for both image platforms.
The [source collector](../scripts/prepare-container-sources.py) verifies authenticated source
metadata and archive hashes; it does not claim to rebuild upstream vendor/OS binaries.

The JAR and launcher archives retain the [repair capsule](../native-provider/README.md), its Maven
identity, matching patched/helper sources and mixed-license attribution. Parent-classpath provider
classes retain their publisher bytes. Runtime SBOM checks include the capsule and the actual
resolved components; source/legal content is inspected separately from component identification.

The GitHub release offers the JAR, ZIP/TAR launchers, owned source archive, container source archive,
source manifest/checksums and `SHA256SUMS`. The data-only source image at
`ghcr.io/resoltico/gridgrind:X.Y.Z-sources` contains both source deliveries:

    docker create --platform linux/amd64 --name gridgrind-sources ghcr.io/resoltico/gridgrind:X.Y.Z-sources /source-only
    docker cp gridgrind-sources:/sources/. ./container-sources
    docker rm gridgrind-sources
    cd container-sources
    sha256sum --check gridgrind-container-sources-X.Y.Z.sha256

The container-source checksum file covers that archive and manifest. Verify the owned project
archive against the release's `SHA256SUMS` as well. On macOS, `shasum -a 256 -c <checksum-file>`
provides the checksum-checking command.

[Project licensing](../PROJECT-NOTICE.md) and [third-party notices](../NOTICE) describe the scopes.
Source acquisition checks expected sizes and space; individual public release assets stay below
2 GiB. The collector's free-space estimate covers source copies and scratch, not an overall disk
quota. Failed preparation retains diagnostics and refuses an occupied output directory.

## Retention

[The pruner](../scripts/prune-container-versions.py) retains the five newest stable semantic-version
roots and matching source tags. Protected subject/fallback metadata and unassociated tagged roots
can retain older graphs indefinitely. This is not a five-graph or physical-storage bound, and it
does not guarantee a merge with concurrent external referrer updates. The pruner verifies graph
hashes and fresh package state before deletion; its standalone default is dry-run. Missing retired
nodes and unrelated untagged records remain untouched, while incomplete protected graphs block
deletion.
