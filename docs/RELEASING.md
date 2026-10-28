# Releasing

Use the Java 27 build toolchain configured in gradle.properties. Local builds use 0.0.0-dev.
A release builds the checked-out vX.Y.Z tag with -Pversion=X.Y.Z; the effective version reaches
JAR metadata, archive names, image labels and verification inputs.

Move the reviewed Unreleased bullets into one exact `## [X.Y.Z] - YYYY-MM-DD` section with a
valid calendar date, leaving an empty Unreleased section above it. Docs-only changes do not
need release-note bullets. Prepare and inspect the changelog-derived notes with:

    ./gradlew prepareReleaseNotes -Pversion=X.Y.Z

The task writes build/release-notes.md and rejects missing, duplicated, empty or invalidly dated
entries. Creating/pushing tags and dispatching publication require explicit release authorization.

## Verification and publication

The tag workflow serializes all releases and retention. It runs the deterministic check graph
and builds versioned archives, then builds both runtime platforms without publication. It compares
the image JARs with the verified standalone JAR and runs the container suite against those exact
local image IDs. `-PcontainerImage=<local-image>` selects this borrowed-image check without
rebuilding or deleting the supplied image.

The workflow enables the containerd image store to preserve typed local manifest descriptors.
It inventories both final filesystems and prepares corresponding source before any
runtime publication. It stages verified assets in a draft, exposes the complete sources and
binaries as a non-latest public prerelease, and publishes the data-only source image first.
Anonymous retrieval must reproduce the source bytes. It then publishes the versioned runtime
index with provenance and SPDX SBOMs, verifies platform-manifest/source bindings, applies safe
retention, advances the image's latest tag and promotes the GitHub release to stable/latest.
There is no atomic transaction across GitHub Releases and GHCR; a failure leaves the verified
prerelease available for inspection rather than silently declaring a stable release.

The manual workflow verifies an existing tag and defaults to dry_run=true. A dry run performs
builds, platform tests and source preparation but skips release/registry mutations and pruning.
Draft assets can be resumed. Public prerelease assets must match exactly; they are not replaced.
An existing stable release stops publication. Inspect remote state after an ambiguous network
failure rather than retagging or blindly overwriting published assets.

Publication controls and their tests come from the commit defining the workflow; product code,
build inputs and source inventories remain on the selected release tag. The optional manual
`verified_run` input names a completed tag-push run that passed all checks and stopped at initial
asset publication. Recovery verifies its repository, workflow, attempt, tag commit and successful
steps, then requires both rebuilt image IDs to equal the previously tested IDs. Only then can
the earlier deterministic and container checks be reused. Rebuilt archives receive focused
distribution checks, image JARs must match the standalone JAR, and source acquisition, retention
and publication checks run fresh. Missing evidence or different images stops recovery.

## Container source delivery

Runtime inputs are owned by [container-runtime.properties](../release/container-runtime.properties)
and the [Dockerfile](../Dockerfile). The build JDK and contributor toolchain remain separate from
the Corretto runtime. Debian snapshot indexes remain authenticated; only historical metadata
expiry checks are disabled.

[prepare-container-sources.py](../scripts/prepare-container-sources.py) maps both image configurations,
JRE commit identities and installed Debian binaries to exact source versions. It verifies the
complete vendor source archive and every original, auxiliary and Debian patch/build archive
named by authenticated source metadata and `.dsc` hashes. It also includes the runtime build
inputs and their legal notices. It does not claim to have rebuilt upstream vendor/OS binaries.

Release assets include the application JAR, ZIP/TAR launchers, a versioned container source
archive, source manifest/checksums and SHA256SUMS. The same source files are offered beside the
runtime at `ghcr.io/resoltico/gridgrind:X.Y.Z-sources`. This is data, not a runnable service:

    docker create --platform linux/amd64 --name gridgrind-sources ghcr.io/resoltico/gridgrind:X.Y.Z-sources /source-only
    docker cp gridgrind-sources:/sources/. ./container-sources
    docker rm gridgrind-sources
    cd container-sources
    sha256sum --check gridgrind-container-sources-X.Y.Z.sha256

The source tag must be publicly retrievable before the runtime and retained with it. Direct
source delivery avoids making a new GridGrind written source offer. [NOTICE](../NOTICE) identifies
the component terms and scope; packaging checks establish retained bytes, not a general legal
opinion about every redistribution scenario.

Source acquisition preflights exact sizes and available space. The release asset must stay below
2 GiB. Its conservative free-space budget covers six possible source copies: upstream payloads,
release archive, BuildKit upload snapshot, source-image layer, downloaded blob and unpacked layer,
plus 512 MiB of scratch headroom. Public verification streams bytes instead of making another
archive copy. Failed source preparation retains diagnostic evidence and refuses an occupied
output directory; review it before resuming.

## Retention

[prune-container-versions.py](../scripts/prune-container-versions.py) retains complete manifest
graphs for the five newest stable runtime versions and their matching source tags. All other
nonretired aliases and source-only incoming versions remain protected. It verifies manifest
hashes and a fresh package-state snapshot before applying a deletion plan; its default is dry-run.
Only resolved manifests in retired graphs can be deleted. Missing retired nodes and unrelated
untagged records are left alone; an incomplete protected graph or other error blocks deletion.
Timestamps choose stable roots, never their child retention boundaries.
