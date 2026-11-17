# Development

GridGrind uses one pinned Eclipse Temurin 27 JDK for build, development and runtime. Start
with the [contributor container](#contributor-container) or the [local toolchain setup](#toolchain-and-dependencies).
Use the checked-in Gradle wrapper. Configuration and lockfiles own the exact inputs; see
[toolchain and dependencies](#toolchain-and-dependencies).

## Checks

Run the normal development gate:

    ./check.sh

The wrapper checks the shell JDK and enters one Gradle graph for check, coverage, the fat JAR
and both launcher distributions. Valid test execution data is reused with unchanged inputs.
Container suites, active fuzzing and product mutation campaigns use explicit tasks. The
included-build tests exercise a bounded real PIT fixture to verify the mutation machinery.

Use focused Gradle tests while changing a behavior. Product tests live beside their owners;
architecture-tests contains ArchUnit rules and negative controls, while engine:parityTest
compares engine behavior with Apache POI within the [declared parity scope](POI_PARITY.md). The golden CLI corpus retains request materialization,
execution and workbook readback through both in-process and packaged-JAR launchers. The
`publicNativeTest` suite separately exercises the public API with unmodified vendor classes
and the packaged repair capsule. `publicNativeLifetimeTest` checks realm reclamation with JVM
soft-reference retention disabled; both are required by `check`. Engine unit tests use explicit
repair compiler/runtime views and do not replace those public-boundary checks.

Re-record the golden corpus with `./gradlew :cli:test --tests '*GoldenCorpusTest' -Pgolden.update=true`,
then review each snapshot diff. Primary outcomes are shared across launchers; case-specific
`.stderr.json` files declare exact standalone versus embedding diagnostics. Generated signature
IDs and previews are normalized only after independent persisted-artifact checks. Authored IDs,
unexpected frames and additional primary output remain visible.

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

    ./gradlew :core:pitest -PmutationClasses=dev.erst.gridgrind.core.vocabulary.RgbHex -PmutationTests=dev.erst.gridgrind.core.vocabulary.RgbHexTest -PmutationMutators=NEGATE_CONDITIONALS --max-workers=1

The comma-separated selectors override the module defaults for that invocation. A narrow run
is not evidence for a complete campaign. Active fuzzing has its own limits in [FUZZING.md](FUZZING.md).

## CI and cache ownership

GitHub runs deterministic checks, product Docker smoke and applicable contributor-container
checks independently; the aggregate Gate requires all applicable results. The deterministic
job does not install Docker build tooling. Gradle cache writes are restricted to the default
branch; other refs restore read-only. Setup-gradle removes unused user-home cache entries
before a permitted cache save, including after a failed build. Failed deterministic, Docker
and contributor-container checks retain their own diagnostic reports for seven days.

Compiler outputs use Gradle's normal up-to-date tracking and stale-class cleanup. The bundled
JAR also stays up to date locally, but is not stored in the build cache: each cached bundle
duplicates its dependencies. Clean or changed bundles are assembled again; compile and test
caching remain enabled. ZIP/TAR packaging uses Gradle's uncached default.

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
execution and native workbook mechanics, with POI confined to unexported native integration packages.
CLI contains commands, rendering and host IO; workspace materialization stays with that host.

[native-provider](../native-provider/README.md) builds the source-pinned provider repair capsule
used by the engine's private native invocation context. The capsule retains repaired classes,
matching sources and legal files under `META-INF/gridgrind/native-provider/`; upstream classes
on the parent classpath remain unchanged. Its Gradle producer supplies explicit compiler views
and the runtime resource artifact, so build the declared project graph rather than replacing
vendor JARs or adding compiled repair directories to the runtime classpath.

JPMS defines module readability and exports. ArchUnit checks package boundaries, native-type
confinement, signatures and closed model shapes against actual compiled product roots.
Build logic is Kotlin in gradle/build-logic. Its tests use ProjectBuilder for task-level contracts
and TestKit for actual Gradle executions.

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

Use the authoritative configuration for exact versions:

- [gradle.properties](../gradle.properties): Java runtime/compiler, bytecode target and local version.
- [Version catalog](../gradle/libs.versions.toml) and module lockfiles: dependencies and build plugins.
- [Wrapper properties](../gradle/wrapper/gradle-wrapper.properties): Gradle distribution and checksum.
- [Runtime inputs](../release/container-runtime.properties) and [Dockerfile](../Dockerfile): one pinned Temurin JDK/Ubuntu image and authenticated package snapshot for every Java role.
- [GitHub workflows](../.github/workflows): action SHAs, hosted runners and CI tool versions.
- [Managed CI JDK setup](../.github/actions/setup-temurin/action.yml): installs the checksum-verified archive selected by the runtime pins and [verifies its full runtime, source and build identities](../scripts/verify-managed-jdk.py).

Docker suites require a working Docker engine and `docker buildx` on PATH. They use anonymous
client configuration while retaining access to the active engine; Docker Desktop is one supported
way to provide that engine locally. Linux CI uses the hosted runner's Docker engine.

Product modules, ordinary tests and deterministic replay use Java 27. The source/bytecode
target is Java 26, matching the supported Kotlin JVM target. PMD supports Java 27. Active JQF fuzzing also runs
on Java 27. Kotlin implements build tasks and convention plugins; their regression tests use Java.

Java formatting uses Spotless's built-in google-java-format integration. Formatting settings
and versions belong to the Java convention plugin and version catalog.

Jackson dependency note: Jackson 3.x databind uses the
`com.fasterxml.jackson.core:jackson-annotations` artifact and package namespace. That is upstream
Jackson design, not a GridGrind version-skew bug.

See [EXECUTION_MODEL.md](./EXECUTION_MODEL.md) for authority and publication,
[CLI.md](./CLI.md) for host commands, [FUZZING.md](./FUZZING.md) for fuzzing,
and [RELEASING.md](./RELEASING.md) for release preparation and publication.

## Contributor container

Contributor configuration is owned by [.devcontainer/devcontainer.json](../.devcontainer/devcontainer.json)
and the `contributor` target in the root [Dockerfile](../Dockerfile). Open the checkout in that
container for the pinned JDK/tooling/font stack. The startup cache repair and opt-in devcontainerTest exercise
caller-writable cache mounts. Dependency update decisions follow
[the automation policy](DEPENDENCY_AUTOMATION_POLICY.md).

All Java container targets derive from the same immutable Eclipse Temurin 27 JDK on Ubuntu 26.04.
[Runtime inputs](../release/container-runtime.properties) retain the full binary, source and
build identities; the shorter upstream image tag is not the full build identity.
The application includes the full JDK. Contributor shell utilities extend the shared base and do
not enter the application target. The contributor target installs the checksum-pinned Docker CLI and Buildx.
The pinned devcontainer feature uses those tools to configure socket access; Compose and the
feature's Buildx installation are disabled.
Editor extensions are recommendations and are not prerequisites for a reproducible build.

For a clean native checkout, install the host-appropriate Temurin JDK with the full
`TEMURIN_VERSION` identity in the [runtime inputs](../release/container-runtime.properties).
Those pins include the managed Linux archives and identify the vendor release; use its matching
macOS archive for a macOS host and verify the publisher checksum. The contributor container
provides the complete pinned Ubuntu environment on either host. Set `JAVA_HOME` to the JDK
installation and prepend `$JAVA_HOME/bin` to `PATH`:

```sh
export JAVA_HOME=/absolute/path/to/the/pinned/jdk
export PATH="$JAVA_HOME/bin:$PATH"
java --version
javac --version
python3 scripts/verify-managed-jdk.py "$JAVA_HOME"
./check.sh
```

The identity helper needs Python 3. Native-provider compilation also verifies the pinned
Temurin runtime identity. The shell precondition reports an actionable error for a missing
compiler or wrong major.
A Docker engine and Buildx are additionally required for the explicit image suites.

Build the contributor target directly with:

    docker build --target contributor -t gridgrind-contributor .

## Raspberry Pi

Use a 64-bit ARM Linux host and a Docker engine with `linux/arm64` support. A 32-bit Raspberry Pi
OS installation cannot run these images. From a clean checkout, build the same contributor target:

    docker buildx build --platform linux/arm64 --load --target contributor -t gridgrind-contributor .

Open the checkout using [.devcontainer/devcontainer.json](../.devcontainer/devcontainer.json),
then run `java --version` and `./check.sh` in its terminal. The container supplies Ubuntu 26.04,
the exact pinned Temurin 27 JDK, fonts and contributor tools. The host does not need a separate
JDK for container development. The devcontainer configures access to the host Docker socket;
container suites still need a working host engine. Its named Gradle/cache mounts are retained
across container recreation.

The application image also supports `linux/arm64`; it uses the same Java/Ubuntu base. Follow
the [container quick start](QUICK_START.md#container) for workspace mounts and host UID/GID.
For a native local build instead, install the matching ARM64 Temurin binary identified by the
[runtime pins](../release/container-runtime.properties) and follow the local setup above.

## Runtime container behavior

The image's default locale is `en_US.UTF-8` and timezone is UTC. Application entry uses Java
without the vendor certificate-import shell. The installed Temurin truststore is retained;
`USE_SYSTEM_CA_CERTS` and `/certificates` are not GridGrind configuration options. Runtime HOME
and font caches live under `/tmp`; read-only-root operation needs an explicit writable `/tmp`
and the host-authorized scratch/work mounts. Host UID/GID overrides remain supported.
