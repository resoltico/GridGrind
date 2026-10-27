# Releasing

Local builds use 0.0.0-dev from gradle.properties. A release builds the checked-out vX.Y.Z tag
with -Pversion=X.Y.Z; the same version reaches JAR metadata, distribution names and the image.

Before requesting a release, add its exact versioned section to CHANGELOG.md and complete the
normal development checks. Prepare and inspect its notes locally with:

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
and pruning. A fork or act can exercise this path without publishing. Creating or pushing tags, dispatching publication workflows and publishing artifacts each
require explicit release authorization.

Dependencies are pinned by the version catalog and Gradle dependency locks. Update lock state
through Gradle's --write-locks option when intentionally changing dependencies; review the result
alongside the dependency change.
