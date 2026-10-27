# Development

GridGrind requires a Java 26 JDK. Use the checked-in Gradle wrapper; gradle.properties owns
the Java pin and local version, and gradle/libs.versions.toml owns dependency versions.
Gradle dependency locks retain the resolved build inputs.

## Checks

Run the normal development gate:

    ./check.sh

The wrapper checks the shell JDK and enters one Gradle graph for check, coverage, the fat JAR
and both launcher distributions. Valid test execution data is reused with unchanged inputs.
It does not run Docker, active fuzzing or mutation campaigns.

Use focused Gradle tests while changing a behavior. Product tests live beside their owners;
architecture-tests contains ArchUnit rules and negative controls, while engine:parityTest
compares engine behavior with Apache POI. The golden CLI corpus retains request materialization,
execution and workbook readback through both in-process and packaged-JAR launchers.

Explicit additional suites are:

    ./gradlew containerTest
    ./gradlew devcontainerTest
    ./gradlew fuzzReplay
    ./gradlew mutationReport

Container suites use anonymous Docker client configuration and the active local engine.
They own their temporary images, containers and volumes. They are excluded from normal test
and coverage prerequisites. Mutation and coverage reports are advisory, without score thresholds.
Active fuzzing is described in [FUZZING.md](./FUZZING.md); ordinary check includes all committed
seed inputs and generator/invariant regressions.

## Campaign resources

PIT uses at most four workers, further limited by available CPUs and Gradle's `--max-workers`.
The core and engine scopes run sequentially; each PIT coordinator and worker has a 512 MiB
heap ceiling. Unchanged inputs reuse the fixed report directory instead of rerunning PIT.
Temporary PIT state is reset only for a real execution, removed on success and retained on
failure until the next run.

Use a narrow class, test and mutator selection when checking the machinery:

    ./gradlew :core:pitest -PmutationClasses=dev.erst.gridgrind.core.vocabulary.RgbHex -PmutationTests=dev.erst.gridgrind.core.vocabulary.RgbHexTest -PmutationMutators=MATH --max-workers=1

The comma-separated selectors override the module defaults for that invocation. A narrow run
is not evidence for a complete campaign. Active fuzzing has its own limits in [FUZZING.md](FUZZING.md).

## CI and cache ownership

GitHub runs deterministic checks, product Docker smoke and applicable contributor-container
checks independently; the aggregate Gate requires all applicable results. The deterministic
job does not install Docker build tooling. Gradle cache writes are restricted to the default
branch; other refs restore read-only. Setup-gradle removes unused user-home cache entries
before a permitted cache save, including after a failed build. Failed check reports are
retained as a seven-day artifact.

The root check finalizer retains three configuration-cache reports: the current input report
and two preceding comparisons. This limits retained diagnostics after checks, not active
writes or arbitrary standalone Gradle invocations. Project test reports and campaign findings
have separate ownership and are not removed by that finalizer.

Gradle owns shared user-home dependency and local build caches; the repository does not impose
a byte quota on that shared directory. The included build reuses one workspace-local TestKit
home across its functional tests. Deleting these caches routinely forces downloads and build
setup again. Preserve useful failure evidence before deliberately cleaning project outputs.

## Architecture

Core contains the pure model, metadata, validation, discovery, grants and executable resolved
plans. JSON owns tolerant intake, binding, schemas and derived wire configuration. Engine owns
execution and native workbook mechanics, with POI confined to its feature/backend packages.
CLI contains commands, rendering and host IO; workspace materialization stays with that host.

JPMS defines module readability and exports. ArchUnit checks package boundaries, native-type
confinement, signatures and closed model shapes against actual compiled product roots.
Build logic is Kotlin in gradle/build-logic; its tests use Gradle's ProjectBuilder.

## Generated source data

Regenerate JSON examples from the pure recipe registry with:

    ./gradlew :cli:writeRepositoryExamples

The committed binary recipe assets are fixed test inputs, including the separately audited
Apache POI Custom XML fixture. Changes to those assets require reviewing their provenance,
NOTICE and packaging checks.

Regenerate exchanged-document schemas with:

    ./gradlew :json:generateSchemas

Reference verification checks committed output without rewriting it. Tests verify examples,
schemas, documentation links and transport behavior. Generate reference pages, limits and
problem descriptions from metadata with:

    ./gradlew :cli:generateReferenceDocs

Use `./gradlew verifyGeneratedDocs` to check committed references and schemas without rewriting
them. Generated workbooks belong in build or test-owned temporary directories.

## Toolchain and dependencies

| Component | Version |
|:----------|:--------|
| Contributor devcontainer | pinned glibc base image plus Azul Zulu 26 JDK; see [contributor configuration](../.devcontainer/devcontainer.json) |
| Java | 26 |
| Docker runtime | Docker Desktop daemon plus `docker buildx` reachable through the active shell `docker` command; smoke and release verification use an anonymous `DOCKER_CONFIG` while still targeting the active local Docker engine |
| Architecture tests | ArchUnit 1.5.0 with its JUnit 6 integration in `architecture-tests:architectureTest` |
| Apache POI | 5.5.1 |
| Jackson Databind | 3.2.2 |
| JUnit Jupiter | 6.1.3 |
| Log4j Core | 2.26.1 |

Product modules and shared build logic target Java 26. Kotlin `2.4.20` implements build tasks
and convention plugins; their regression tests use Java. The included build reads its Java
target from the root gradle.properties pin.

Jackson dependency note: Jackson 3.x databind intentionally still uses the
`com.fasterxml.jackson.core:jackson-annotations` artifact and package namespace. That is upstream
Jackson design, not a GridGrind version-skew bug.


See [EXECUTION_MODEL.md](./EXECUTION_MODEL.md) for authority and publication,
[CLI.md](./CLI.md) for host commands, [FUZZING.md](./FUZZING.md) for fuzzing,
and [RELEASING.md](./RELEASING.md) for the tag-driven release workflow.

## Contributor container

Contributor configuration is owned by [.devcontainer/devcontainer.json](../.devcontainer/devcontainer.json)
and its [Dockerfile](../.devcontainer/Dockerfile). Open the checkout in that container for the
pinned JDK/tooling/font stack. The startup cache repair and opt-in devcontainerTest exercise
caller-writable cache mounts. Dependency update decisions follow
[the automation policy](DEPENDENCY_AUTOMATION_POLICY.md).
