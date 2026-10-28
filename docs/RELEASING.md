# Releasing

Use the Java 27 toolchain/runtime configured in gradle.properties.

Local builds use 0.0.0-dev from gradle.properties. A release builds the checked-out vX.Y.Z tag
with -Pversion=X.Y.Z; the same version reaches JAR metadata, distribution names and the image.

Prepare the release summary from CHANGELOG.md: move the reviewed Unreleased bullets into the
exact versioned, dated section and leave an empty Unreleased section above it. Docs-only changes
do not need release-note bullets. Complete the normal development checks before publishing.
Prepare and inspect the notes locally with:

```bash
./gradlew prepareReleaseNotes -Pversion=X.Y.Z
```

The task writes build/release-notes.md. Its heading must be exactly `## [X.Y.Z] - YYYY-MM-DD`
with a valid calendar date. A missing, duplicated or empty version section fails.
The development version cannot be published through this preparation step.

The release workflow runs check.sh, then the opt-in container tests, before creating a GitHub
release containing the fat JAR and both launcher distributions. It publishes linux/amd64 and
linux/arm64 images under the version and latest tags, with provenance and an SBOM. The
prune job retains complete package-version groups for the five newest stable releases.

The manual workflow accepts an existing tag's version and defaults to dry_run=true. Dry runs
perform the build and verification but skip release creation, registry login, image publication
and pruning. Creating or pushing tags, dispatching publication workflows and publishing artifacts
each require explicit release authorization.

Dependencies are pinned by the version catalog and Gradle dependency locks. Update lock state
through Gradle's --write-locks option when intentionally changing dependencies; review the result
alongside the dependency change.

Review [NOTICE](../NOTICE) for the delivered runtime, shaded components and container terms.
The packaging checks verify retained legal bytes and selected content boundaries. They do not
establish every source-delivery obligation for a particular redistribution. Before container
publication, establish the corresponding-source arrangements for its actual JRE and OS packages
and distribution method; retained upstream offers and source links do not create a GridGrind offer.
