# Private dependency tuples

The [version catalog](../gradle/libs.versions.toml) owns authored dependency pins. Product
[module lockfiles](../core/gradle.lockfile), the [root lockfile](../gradle.lockfile) and the
[included-build lockfile](../gradle/build-logic/gradle.lockfile) retain exact resolved
configurations. Consult those files for current versions; this guide records ownership and
compatibility reasons rather than a second version inventory.

Application libraries use scoped alignment constraints. Compiler, formatter, PMD and PIT
internals follow their owning tool's declared tuple. A newer standalone component release is
not evidence that replacing a tool-private dependency preserves its compiler, parser or
reporting API. Review the publisher metadata and exercise that owner's path before changing it.

| Components | Owner and retained relationship |
| --- | --- |
| Durian SWT OS, OkHttp and Okio | Spotless's Solstice provisioning subtree owns these transport/provisioning APIs. |
| Error Prone dataflow | Error Prone owns its compiler-analysis implementation. |
| NullAway dataflow and checker-qual | NullAway owns its analysis framework tuple. |
| PMD checker-qual and PCollections | PMD owns its analysis dependencies; Error Prone can independently require a different PCollections family. |
| Saxon-HE and XMLResolver | PMD owns the XPath engine and resolver API pair. |
| JUL-to-SLF4J | PMD's worker bridge follows PMD; application SLF4J belongs to a separate graph. |
| jopt-simple, StringTemplate and ANTLR | PIT owns its command parser and HTML-report implementation. |
| Google Java Format | Error Prone's compiler-private formatter API is separate from Spotless's public integration. |
| CommonMark and GFM tables | The selected Google Java Format integration owns this parser/extension API pair. |
| Kotlin compiler, reflect, coroutines and annotations | Gradle's embedded Kotlin and Kotlin-plugin metadata select internal subtrees, including private ABI validation. These need not match every public library release. |

[FormatterDependencies](../gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/FormatterDependencies.kt)
aligns public Guava and annotation libraries only in the selected detached Spotless formatter
graph. It leaves Error Prone's compiler-private formatter API unchanged. The
[included build](../gradle/build-logic/build.gradle.kts) declares build-plugin dependencies
and their scoped constraints.

The normal gate exercises Error Prone/NullAway, PMD, formatter provisioning, configuration-cache
reuse, schema validation and deterministic fuzz replay. The build-logic mutation fixture runs a
bounded real PIT campaign and checks report generation. Full mutation and active fuzz campaigns
remain explicit tasks. See [development](DEVELOPER.md) and the
[dependency automation policy](DEPENDENCY_AUTOMATION_POLICY.md).
