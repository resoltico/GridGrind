package dev.erst.gridgrind.buildlogic

import java.math.BigDecimal
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.artifacts.VersionCatalog
import org.gradle.api.artifacts.VersionCatalogsExtension
import org.gradle.api.file.Directory
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.plugins.quality.Pmd
import org.gradle.api.tasks.JavaExec
import org.gradle.api.tasks.SourceSetContainer
import org.gradle.api.tasks.testing.Test
import org.gradle.jvm.toolchain.JavaLanguageVersion
import org.gradle.kotlin.dsl.getByType
import org.gradle.kotlin.dsl.named
import org.gradle.kotlin.dsl.register
import org.gradle.process.JavaForkOptions
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport

class GridGrindJazzerConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        with(project) {
            pluginManager.apply("java")
            pluginManager.apply("gridgrind.java-conventions")

            description = "Local-only Jazzer fuzzing layer for GridGrind"

            val topology = JazzerTopology.load(this)
            val libs = extensions.getByType<VersionCatalogsExtension>().named("libs")
            val repositoryLayout = GridGrindRepositoryLayout.locate(this)
            val jazzerMainPmdRuleset = repositoryLayout.repositoryRoot.resolve("gradle/pmd/jazzer-ruleset.xml")
            val jazzerFuzzPmdRuleset = repositoryLayout.repositoryRoot.resolve("gradle/pmd/jazzer-fuzz-ruleset.xml")
            val gridgrindJavaVersion =
                providers.gradleProperty("gridgrindJavaVersion").map(String::toInt).get()
            val jazzerMaxDuration = providers.gradleProperty("jazzerMaxDuration").orNull
            val jazzerMaxExecutions = providers.gradleProperty("jazzerMaxExecutions").orNull

            val javaExtension = extensions.getByType<JavaPluginExtension>()
            javaExtension.toolchain.languageVersion.set(JavaLanguageVersion.of(gridgrindJavaVersion))

            val sourceSets = extensions.getByType<SourceSetContainer>()
            val mainSourceSet = sourceSets.getByName("main")
            val fuzzSourceSet = sourceSets.create("fuzz") {
                java.setSrcDirs(listOf("src/fuzz/java"))
                resources.setSrcDirs(listOf("src/fuzz/resources"))
            }
            fuzzSourceSet.compileClasspath += mainSourceSet.output
            fuzzSourceSet.runtimeClasspath += mainSourceSet.output

            configurations.named(fuzzSourceSet.implementationConfigurationName) {
                extendsFrom(configurations.getByName("implementation"))
            }
            configurations.named(fuzzSourceSet.runtimeOnlyConfigurationName) {
                extendsFrom(configurations.getByName("runtimeOnly"))
            }

            dependencies.apply {
                add("implementation", platform(libs.library("junit-bom")))
                add("implementation", libs.library("junit-platform-launcher"))
                add("implementation", libs.library("jazzer-api"))
                add("implementation", libs.library("jazzer"))
                add("implementation", libs.library("jazzer-junit"))
                add("implementation", libs.library("poi-ooxml"))
                add("implementation", libs.library("jackson-databind"))
                add("implementation", "dev.erst.gridgrind:protocol")
                add("implementation", "dev.erst.gridgrind:engine")
                add("runtimeOnly", libs.library("log4j-core"))

                add("testImplementation", platform(libs.library("junit-bom")))
                add("testImplementation", libs.library("junit-jupiter"))
                add("testImplementation", libs.library("jazzer-junit"))
                add("testImplementation", libs.library("poi-ooxml"))
                add("testImplementation", libs.library("jackson-databind"))
                add("testImplementation", "dev.erst.gridgrind:protocol")
                add("testImplementation", "dev.erst.gridgrind:engine")
                add("testRuntimeOnly", libs.library("junit-platform-launcher"))

                add(fuzzSourceSet.implementationConfigurationName, platform(libs.library("junit-bom")))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("junit-jupiter"))
                add(fuzzSourceSet.runtimeOnlyConfigurationName, libs.library("junit-platform-launcher"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("jazzer-junit"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("jazzer-api"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("poi-ooxml"))
                add(fuzzSourceSet.implementationConfigurationName, "dev.erst.gridgrind:protocol")
                add(fuzzSourceSet.implementationConfigurationName, "dev.erst.gridgrind:engine")
            }

            fun JavaExec.configureHarnessRuntime() {
                classpath = fuzzSourceSet.runtimeClasspath
                mainClass.set("dev.erst.gridgrind.jazzer.tool.JazzerHarnessRunner")
                outputs.upToDateWhen { false }
                workingDir = layout.projectDirectory.asFile
                enableNativeAccess()
                if (jazzerMaxDuration != null) {
                    systemProperty("jazzer.max_duration", jazzerMaxDuration)
                }
                if (jazzerMaxExecutions != null) {
                    systemProperty("jazzer.max_executions", jazzerMaxExecutions)
                }
            }

            fun JavaExec.configureMainSourceSet() {
                classpath = mainSourceSet.runtimeClasspath
                mainClass.set("dev.erst.gridgrind.jazzer.tool.JazzerCli")
                workingDir = layout.projectDirectory.asFile
                enableNativeAccess()
            }

            fun requiredGradleProperty(name: String): String =
                providers.gradleProperty(name).orNull
                    ?: throw IllegalArgumentException("Missing Gradle property: $name")

            fun registerToolTask(
                name: String,
                descriptionText: String,
                argumentsProvider: () -> List<String>,
            ) = tasks.register<JavaExec>(name) {
                description = descriptionText
                group = "verification"
                configureMainSourceSet()
                notCompatibleWithConfigurationCache(
                    "Local-only Jazzer tool tasks assemble command arguments at execution time.",
                )
                doFirst {
                    setArgs(argumentsProvider())
                }
            }

            fun registerFuzzTask(
                target: JazzerRunTargetSpec,
                harness: JazzerHarnessSpec,
            ) = tasks.register<JavaExec>(target.taskName) {
                description = "Actively fuzzes ${target.displayName.lowercase()}."
                group = "verification"
                configureHarnessRuntime()
                args("--class", harness.className)
                workingDir = runDirectory(layout.projectDirectory.dir(target.workingDirectory))
                doFirst {
                    workingDir.mkdirs()
                }
                environment("JAZZER_FUZZ", "1")
            }

            fun registerRegressionTask(harness: JazzerHarnessSpec) =
                tasks.register<JavaExec>("regression${harness.key.toTaskSuffix()}") {
                    description = "Replays committed ${harness.key} Jazzer inputs in regression mode."
                    group = "verification"
                    configureMainSourceSet()
                    mainClass.set("dev.erst.gridgrind.jazzer.tool.JazzerRegressionRunner")
                    args("--target", harness.key)
                    workingDir = layout.projectDirectory.asFile
                }

            val regressionTasks =
                topology.harnesses.map(::registerRegressionTask)

            val fuzzTasks =
                topology.runTargets
                    .filter(JazzerRunTargetSpec::activeFuzzing)
                    .map { target ->
                        registerFuzzTask(target = target, harness = topology.harness(target.key))
                    }

            val regressionTarget = topology.runTarget("regression")
            val jazzerRegression =
                tasks.register(regressionTarget.taskName) {
                    description = "Runs all Jazzer harnesses in regression mode."
                    group = "verification"
                    dependsOn(regressionTasks)
                }

            tasks.named<Test>("test") {
                val supportTestClassCount =
                    fileTree("src/test/java") {
                        include("**/*Test.java")
                    }.files.size
                description = "Runs deterministic Jazzer support tests."
                group = "verification"
                useJUnitPlatform()
                maxParallelForks = 1
                enableNativeAccess()
                doFirst {
                    addTestListener(JazzerSupportTestPulseListener(supportTestClassCount))
                }
            }

            tasks.withType(Pmd::class.java)
                .matching { task -> task.name == "pmdMain" }
                .configureEach {
                    ruleSetFiles = files(jazzerMainPmdRuleset)
                    ruleSets = emptyList()
                }

            tasks.withType(Pmd::class.java)
                .matching { task -> task.name == "pmdFuzz" }
                .configureEach {
                    ruleSetFiles = files(jazzerFuzzPmdRuleset)
                    ruleSets = emptyList()
                }

            val jazzerCoverageClasses =
                mainSourceSet.output.classesDirs.asFileTree.matching {
                    exclude(*JAZZER_COVERAGE_EXCLUSIONS.toTypedArray())
                }

            tasks.named<JacocoReport>("jacocoTestReport") {
                classDirectories.setFrom(jazzerCoverageClasses)
            }

            tasks.named<JacocoCoverageVerification>("jacocoTestCoverageVerification") {
                enabled = false
            }

            val jazzerCoverageVerification =
                tasks.register<JacocoCoverageVerification>("jazzerCoverageVerification") {
                    description =
                        "Verifies deterministic-unit coverage for the Jazzer support-contract subset."
                    group = "verification"
                    dependsOn(tasks.named<Test>("test"))
                    executionData.from(layout.buildDirectory.file("jacoco/test.exec"))
                    sourceDirectories.setFrom(mainSourceSet.allJava.srcDirs)
                    classDirectories.setFrom(jazzerCoverageClasses)
                    violationRules {
                        rule {
                            limit {
                                minimum = BigDecimal(JAZZER_COVERAGE_MINIMUM)
                            }
                        }
                    }
                }

            tasks.named("check") {
                dependsOn(jazzerCoverageVerification)
            }

            tasks.named("check") {
                dependsOn(jazzerRegression)
            }

            tasks.register("fuzzAllLocal") {
                description = "Runs all active local-only fuzzing tasks."
                group = "verification"
                dependsOn(fuzzTasks)
            }

            fuzzTasks.windowed(size = 2, step = 1, partialWindows = false).forEach { (first, second) ->
                second.configure {
                    mustRunAfter(first)
                }
            }

            val jazzerSupportTests = tasks.named<Test>("test")
            regressionTasks.windowed(size = 2, step = 1, partialWindows = false).forEach { (first, second) ->
                second.configure {
                    mustRunAfter(first)
                    mustRunAfter(jazzerSupportTests)
                }
            }
            regressionTasks.firstOrNull()?.configure {
                mustRunAfter(jazzerSupportTests)
            }

            tasks.register<CleanLocalCorpusTask>("cleanLocalCorpus") {
                description = "Deletes generated Jazzer corpora under .local."
                group = "build"
                localDirectory.set(layout.projectDirectory.dir(".local"))
            }

            tasks.register<CleanLocalFindingsTask>("cleanLocalFindings") {
                description = "Deletes local crash files and non-corpus run state under .local."
                group = "build"
                runsDirectory.set(layout.projectDirectory.dir(".local/runs"))
            }

            registerToolTask(
                "jazzerSummarizeRun",
                "Builds the latest summary artifacts for one completed Jazzer run.",
            ) {
                listOf(
                    "summarize-run",
                    "--project-dir",
                    layout.projectDirectory.asFile.absolutePath,
                    "--target",
                    requiredGradleProperty("jazzerTarget"),
                    "--task",
                    requiredGradleProperty("jazzerTaskName"),
                    "--log",
                    requiredGradleProperty("jazzerLogPath"),
                    "--history",
                    requiredGradleProperty("jazzerHistoryDirectory"),
                    "--started-at",
                    requiredGradleProperty("jazzerStartedAt"),
                    "--finished-at",
                    requiredGradleProperty("jazzerFinishedAt"),
                    "--exit-code",
                    requiredGradleProperty("jazzerExitCode"),
                    "--corpus-before-files",
                    requiredGradleProperty("jazzerCorpusBeforeFiles"),
                    "--corpus-before-bytes",
                    requiredGradleProperty("jazzerCorpusBeforeBytes"),
                )
            }

            registerToolTask("jazzerStatus", "Shows a concise latest-summary view across Jazzer targets.") {
                buildList {
                    add("status")
                    add("--project-dir")
                    add(layout.projectDirectory.asFile.absolutePath)
                    providers.gradleProperty("jazzerTarget").orNull?.let {
                        add("--target")
                        add(it)
                    }
                }
            }

            registerToolTask("jazzerReport", "Shows the detailed latest-summary report for Jazzer targets.") {
                buildList {
                    add("report")
                    add("--project-dir")
                    add(layout.projectDirectory.asFile.absolutePath)
                    providers.gradleProperty("jazzerTarget").orNull?.let {
                        add("--target")
                        add(it)
                    }
                }
            }

            registerToolTask("jazzerListFindings", "Lists replayed local finding artifacts.") {
                buildList {
                    add("list-findings")
                    add("--project-dir")
                    add(layout.projectDirectory.asFile.absolutePath)
                    providers.gradleProperty("jazzerTarget").orNull?.let {
                        add("--target")
                        add(it)
                    }
                }
            }

            registerToolTask("jazzerListCorpus", "Lists local corpus state and promoted inputs.") {
                buildList {
                    add("list-corpus")
                    add("--project-dir")
                    add(layout.projectDirectory.asFile.absolutePath)
                    providers.gradleProperty("jazzerTarget").orNull?.let {
                        add("--target")
                        add(it)
                    }
                }
            }

            registerToolTask("jazzerReplay", "Replays one local input against a single Jazzer harness.") {
                buildList {
                    add("replay")
                    add("--project-dir")
                    add(layout.projectDirectory.asFile.absolutePath)
                    add("--target")
                    add(requiredGradleProperty("jazzerTarget"))
                    add("--input")
                    add(requiredGradleProperty("jazzerInput"))
                    if (providers.gradleProperty("jazzerJsonOutput").orNull == "true") {
                        add("--json")
                    }
                }
            }

            registerToolTask("jazzerPromote", "Promotes one local input into committed Jazzer regression resources.") {
                listOf(
                    "promote",
                    "--project-dir",
                    layout.projectDirectory.asFile.absolutePath,
                    "--target",
                    requiredGradleProperty("jazzerTarget"),
                    "--input",
                    requiredGradleProperty("jazzerInput"),
                    "--name",
                    requiredGradleProperty("jazzerName"),
                )
            }

            registerToolTask(
                "jazzerRefreshPromotedMetadata",
                "Refreshes promoted Jazzer metadata and replay text from the current replay contract.",
            ) {
                buildList {
                    add("refresh-promoted-metadata")
                    add("--project-dir")
                    add(layout.projectDirectory.asFile.absolutePath)
                    providers.gradleProperty("jazzerTarget").orNull?.let {
                        add("--target")
                        add(it)
                    }
                }
            }
        }
    }

    private fun VersionCatalog.library(name: String): Any =
        findLibrary(name).orElseThrow { IllegalArgumentException("Missing version-catalog library: $name") }.get()

    private fun JavaForkOptions.enableNativeAccess() {
        jvmArgs(NATIVE_ACCESS_ARGUMENT)
    }

    private fun runDirectory(directory: Directory) = directory.asFile

    private fun String.toTaskSuffix(): String =
        split('-').joinToString(separator = "") { segment ->
            segment.replaceFirstChar { character -> character.titlecase() }
        }

    private companion object {
        private const val NATIVE_ACCESS_ARGUMENT = "--enable-native-access=ALL-UNNAMED"
        private const val JAZZER_COVERAGE_MINIMUM = "0.72"
        private val JAZZER_COVERAGE_EXCLUSIONS =
            listOf(
                "dev/erst/gridgrind/jazzer/support/FuzzDataDecoders*.class",
                "dev/erst/gridgrind/jazzer/support/GeneratedProtocolWorkflow*.class",
                "dev/erst/gridgrind/jazzer/support/HarnessTelemetry*.class",
                "dev/erst/gridgrind/jazzer/support/JazzerGridGrindFuzzData*.class",
                "dev/erst/gridgrind/jazzer/support/OperationSequenceModel*.class",
                "dev/erst/gridgrind/jazzer/support/StyleKindIntrospection*.class",
                "dev/erst/gridgrind/jazzer/support/WorkbookStyleInputs*.class",
                "dev/erst/gridgrind/jazzer/tool/JazzerCli*.class",
                "dev/erst/gridgrind/jazzer/tool/JazzerReportSupport*.class",
                "dev/erst/gridgrind/jazzer/tool/RunMetrics*.class",
            )
    }
}
