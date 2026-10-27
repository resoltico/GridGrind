# Gradle

The root build contains core, json, engine, cli and architecture-tests, plus the fuzz source
set. Shared conventions and build tasks are Kotlin in gradle/build-logic.

Use [DEVELOPER.md](./DEVELOPER.md) for the supported commands and ownership boundaries.
The root check.sh invokes one Gradle graph, so tests, reports and delivery archives share
applicable outputs. Docker suites and active campaigns are explicit operations.

Dependency versions come from gradle/libs.versions.toml and Java/version settings from
gradle.properties. Commit intentional Gradle dependency-lock updates with their source changes.
The included build consumes the same Java target and has its own dependency locks.

The root fuzzReplay task uses ordinary JUnit regression replay; fuzz launches the pinned
standalone driver only when explicitly selected. See [FUZZING.md](./FUZZING.md).

CLI distributions are built with shadowDistZip, shadowDistTar and installShadowDist. Their
archive/legal checks and installed-launcher execution run in normal CLI tests. Container tests
are tagged and excluded from normal coverage. Release notes are prepared by prepareReleaseNotes
from the exact versioned CHANGELOG section; see [RELEASING.md](./RELEASING.md).
