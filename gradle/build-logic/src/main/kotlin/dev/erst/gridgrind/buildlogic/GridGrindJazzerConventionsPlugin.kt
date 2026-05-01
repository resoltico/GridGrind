package dev.erst.gridgrind.buildlogic

import java.math.BigDecimal
import javax.inject.Inject
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.artifacts.VersionCatalog
import org.gradle.api.artifacts.VersionCatalogsExtension
import org.gradle.api.file.Directory
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.plugins.quality.Pmd
import org.gradle.api.provider.Property
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.JavaExec
import org.gradle.api.tasks.Optional
import org.gradle.api.tasks.SourceSetContainer
import org.gradle.api.tasks.bundling.Jar
import org.gradle.api.tasks.testing.Test
import org.gradle.jvm.toolchain.JavaLanguageVersion
import org.gradle.kotlin.dsl.getByType
import org.gradle.kotlin.dsl.named
import org.gradle.kotlin.dsl.newInstance
import org.gradle.kotlin.dsl.register
import org.gradle.process.CommandLineArgumentProvider
import org.gradle.process.JavaForkOptions
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport
import org.gradle.testing.jacoco.tasks.rules.JacocoLimit
import org.gradle.testing.jacoco.tasks.rules.JacocoViolationRule
import org.gradle.testing.jacoco.tasks.rules.JacocoViolationRulesContainer

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
            val testSourceSet = sourceSets.getByName("test")
            val fuzzSourceSet = sourceSets.create("fuzz") { fuzzSourceSet ->
                fuzzSourceSet.java.setSrcDirs(listOf("src/fuzz/java"))
                fuzzSourceSet.resources.setSrcDirs(listOf("src/fuzz/resources"))
            }
            val jazzerAgentJar =
                tasks.named<Jar>("jar") {
                    manifest.attributes(
                        mapOf(
                            "Premain-Class" to JAZZER_PREMAIN_CLASS,
                            "Agent-Class" to JAZZER_PREMAIN_CLASS,
                            "Can-Redefine-Classes" to "true",
                            "Can-Retransform-Classes" to "true",
                            "Can-Set-Native-Method-Prefix" to "true",
                        ),
                    )
                }
            fuzzSourceSet.compileClasspath += mainSourceSet.output
            fuzzSourceSet.runtimeClasspath += mainSourceSet.output

            configurations.named(fuzzSourceSet.implementationConfigurationName) { configuration ->
                configuration.extendsFrom(configurations.getByName("implementation"))
            }
            configurations.named(fuzzSourceSet.runtimeOnlyConfigurationName) { configuration ->
                configuration.extendsFrom(configurations.getByName("runtimeOnly"))
            }

            dependencies.apply {
                add("implementation", platform(libs.library("junit-bom")))
                add("implementation", libs.library("junit-platform-launcher"))
                add("implementation", libs.library("jazzer-api"))
                add("implementation", libs.library("jazzer"))
                add("implementation", libs.library("jazzer-junit"))
                add("implementation", libs.library("poi-ooxml"))
                add("implementation", libs.library("poi-ooxml-full"))
                add("implementation", libs.library("jackson-databind"))
                add("implementation", libs.library("xmlsec"))
                add("implementation", libs.library("bcpkix"))
                add("implementation", libs.library("bcprov"))
                add("implementation", libs.library("bcutil"))
                add("implementation", "dev.erst.gridgrind:executor")
                add("implementation", "dev.erst.gridgrind:engine")
                add("runtimeOnly", libs.library("log4j-core"))
                add("runtimeOnly", libs.library("log4j-slf4j2-impl"))

                add("testImplementation", platform(libs.library("junit-bom")))
                add("testImplementation", libs.library("junit-jupiter"))
                add("testImplementation", libs.library("jazzer-junit"))
                add("testImplementation", libs.library("poi-ooxml"))
                add("testImplementation", libs.library("poi-ooxml-full"))
                add("testImplementation", libs.library("jackson-databind"))
                add("testImplementation", libs.library("xmlsec"))
                add("testImplementation", libs.library("bcpkix"))
                add("testImplementation", libs.library("bcprov"))
                add("testImplementation", libs.library("bcutil"))
                add("testImplementation", "dev.erst.gridgrind:executor")
                add("testImplementation", "dev.erst.gridgrind:engine")
                add("testRuntimeOnly", libs.library("junit-platform-launcher"))
                add("testRuntimeOnly", libs.library("log4j-slf4j2-impl"))

                add(fuzzSourceSet.implementationConfigurationName, platform(libs.library("junit-bom")))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("junit-jupiter"))
                add(fuzzSourceSet.runtimeOnlyConfigurationName, libs.library("junit-platform-launcher"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("jazzer-junit"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("jazzer-api"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("poi-ooxml"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("poi-ooxml-full"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("xmlsec"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("bcpkix"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("bcprov"))
                add(fuzzSourceSet.implementationConfigurationName, libs.library("bcutil"))
                add(fuzzSourceSet.implementationConfigurationName, "dev.erst.gridgrind:executor")
                add(fuzzSourceSet.implementationConfigurationName, "dev.erst.gridgrind:engine")
            }

            fun JavaExec.configureHarnessRuntime() {
                classpath = fuzzSourceSet.runtimeClasspath
                mainClass.set("dev.erst.gridgrind.jazzer.tool.JazzerHarnessRunner")
                outputs.upToDateWhen { false }
                workingDir = layout.projectDirectory.asFile
                dependsOn(jazzerAgentJar)
                enableNativeAccess()
                jvmArgs("-javaagent:${jazzerAgentJar.flatMap { it.archiveFile }.get().asFile.absolutePath}")
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

            fun jazzerToolArgumentsProvider(
                operationName: String,
                configuration: JazzerToolArgumentsProvider.() -> Unit = {},
            ): JazzerToolArgumentsProvider =
                objects.newInstance<JazzerToolArgumentsProvider>().apply {
                    operation.set(operationName)
                    projectDirectoryPath.set(layout.projectDirectory.asFile.absolutePath)
                    jsonOutput.convention(false)
                    configuration()
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
                tasks.register(regressionTarget.taskName) { jazzerRegressionTask ->
                    jazzerRegressionTask.description = "Runs all Jazzer harnesses in regression mode."
                    jazzerRegressionTask.group = "verification"
                    jazzerRegressionTask.dependsOn(regressionTasks)
                }

            tasks.named<Test>("test") {
                val supportTestClassCount =
                    fileTree("src/test/java") {
                        include("**/*Test.java")
                    }.files.size
                description = "Runs deterministic Jazzer support tests."
                group = "verification"
                testClassesDirs = testSourceSet.output.classesDirs
                classpath = testSourceSet.runtimeClasspath
                isScanForTestClasses = false
                include("**/*Test.class")
                useJUnitPlatform()
                maxParallelForks = 1
                enableNativeAccess()
                doFirst {
                    addTestListener(JazzerSupportTestPulseListener(supportTestClassCount))
                }
            }

            tasks.withType(Pmd::class.java)
                .matching { task -> task.name == "pmdMain" }
                .configureEach { pmd ->
                    pmd.ruleSetFiles = files(jazzerMainPmdRuleset)
                    pmd.ruleSets = emptyList()
                }

            tasks.withType(Pmd::class.java)
                .matching { task -> task.name == "pmdFuzz" }
                .configureEach { pmd ->
                    pmd.ruleSetFiles = files(jazzerFuzzPmdRuleset)
                    pmd.ruleSets = emptyList()
                }

            val jazzerCoverageClasses =
                mainSourceSet.output.classesDirs.asFileTree.matching { patternFilterable ->
                    patternFilterable.exclude(*JAZZER_COVERAGE_EXCLUSIONS.toTypedArray())
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
                    val jazzerTest = tasks.named<Test>("test")
                    dependsOn(jazzerTest)
                    executionData.from(
                        provider {
                            jazzerTest
                                .get()
                                .extensions
                                .getByType(JacocoTaskExtension::class.java)
                                .destinationFile
                        },
                    )
                    sourceDirectories.setFrom(mainSourceSet.allJava.srcDirs)
                    classDirectories.setFrom(jazzerCoverageClasses)
                    violationRules { rules: JacocoViolationRulesContainer ->
                        rules.rule { rule: JacocoViolationRule ->
                            rule.limit { limit: JacocoLimit ->
                                limit.minimum = BigDecimal(JAZZER_COVERAGE_MINIMUM)
                            }
                        }
                    }
                }

            tasks.named("check") { checkTask ->
                checkTask.dependsOn(jazzerCoverageVerification)
            }

            tasks.named("check") { checkTask ->
                checkTask.dependsOn(jazzerRegression)
            }

            tasks.register("fuzzAllLocal") { fuzzAllTask ->
                fuzzAllTask.description = "Runs all active local-only fuzzing tasks."
                fuzzAllTask.group = "verification"
                fuzzAllTask.dependsOn(fuzzTasks)
            }

            fuzzTasks.windowed(size = 2, step = 1, partialWindows = false).forEach { (first, second) ->
                second.configure { task ->
                    task.mustRunAfter(first)
                }
            }

            val jazzerSupportTests = tasks.named<Test>("test")
            regressionTasks.windowed(size = 2, step = 1, partialWindows = false).forEach { (first, second) ->
                second.configure { task ->
                    task.mustRunAfter(first)
                    task.mustRunAfter(jazzerSupportTests)
                }
            }
            regressionTasks.firstOrNull()?.configure { task ->
                task.mustRunAfter(jazzerSupportTests)
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

            tasks.register<JavaExec>("jazzerSummarizeRun") {
                description = "Builds the latest summary artifacts for one completed Jazzer run."
                group = "verification"
                configureMainSourceSet()
                argumentProviders.add(
                    jazzerToolArgumentsProvider("summarize-run") {
                        target.set(providers.gradleProperty("jazzerTarget"))
                        summarizedTaskName.set(providers.gradleProperty("jazzerTaskName"))
                        logPath.set(providers.gradleProperty("jazzerLogPath"))
                        historyDirectory.set(providers.gradleProperty("jazzerHistoryDirectory"))
                        startedAt.set(providers.gradleProperty("jazzerStartedAt"))
                        finishedAt.set(providers.gradleProperty("jazzerFinishedAt"))
                        exitCode.set(providers.gradleProperty("jazzerExitCode"))
                        corpusBeforeFiles.set(providers.gradleProperty("jazzerCorpusBeforeFiles"))
                        corpusBeforeBytes.set(providers.gradleProperty("jazzerCorpusBeforeBytes"))
                    },
                )
            }

            tasks.register<JavaExec>("jazzerStatus") {
                description = "Shows a concise latest-summary view across Jazzer targets."
                group = "verification"
                configureMainSourceSet()
                argumentProviders.add(
                    jazzerToolArgumentsProvider("status") {
                        target.set(providers.gradleProperty("jazzerTarget"))
                    },
                )
            }

            tasks.register<JavaExec>("jazzerReport") {
                description = "Shows the detailed latest-summary report for Jazzer targets."
                group = "verification"
                configureMainSourceSet()
                argumentProviders.add(
                    jazzerToolArgumentsProvider("report") {
                        target.set(providers.gradleProperty("jazzerTarget"))
                    },
                )
            }

            tasks.register<JavaExec>("jazzerListFindings") {
                description = "Lists replayed local finding artifacts."
                group = "verification"
                configureMainSourceSet()
                argumentProviders.add(
                    jazzerToolArgumentsProvider("list-findings") {
                        target.set(providers.gradleProperty("jazzerTarget"))
                    },
                )
            }

            tasks.register<JavaExec>("jazzerListCorpus") {
                description = "Lists local corpus state and promoted inputs."
                group = "verification"
                configureMainSourceSet()
                argumentProviders.add(
                    jazzerToolArgumentsProvider("list-corpus") {
                        target.set(providers.gradleProperty("jazzerTarget"))
                    },
                )
            }

            tasks.register<JavaExec>("jazzerReplay") {
                description = "Replays one local input against a single Jazzer harness."
                group = "verification"
                configureMainSourceSet()
                argumentProviders.add(
                    jazzerToolArgumentsProvider("replay") {
                        target.set(providers.gradleProperty("jazzerTarget"))
                        inputPath.set(providers.gradleProperty("jazzerInput"))
                        jsonOutput.set(
                            providers.gradleProperty("jazzerJsonOutput").map { it == "true" }.orElse(false),
                        )
                    },
                )
            }

            tasks.register<JavaExec>("jazzerPromote") {
                description = "Promotes one local input into committed Jazzer regression resources."
                group = "verification"
                configureMainSourceSet()
                argumentProviders.add(
                    jazzerToolArgumentsProvider("promote") {
                        target.set(providers.gradleProperty("jazzerTarget"))
                        inputPath.set(providers.gradleProperty("jazzerInput"))
                        promotedName.set(providers.gradleProperty("jazzerName"))
                    },
                )
            }

            tasks.register<JavaExec>("jazzerRefreshPromotedMetadata") {
                description = "Refreshes promoted Jazzer metadata and replay text from the current replay contract."
                group = "verification"
                configureMainSourceSet()
                argumentProviders.add(
                    jazzerToolArgumentsProvider("refresh-promoted-metadata") {
                        target.set(providers.gradleProperty("jazzerTarget"))
                    },
                )
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
        private const val JAZZER_PREMAIN_CLASS = "dev.erst.gridgrind.jazzer.tool.JazzerPremainAgent"
        private const val JAZZER_COVERAGE_MINIMUM = "0.72"
        private val JAZZER_COVERAGE_EXCLUSIONS =
            listOf(
                "dev/erst/gridgrind/jazzer/support/FuzzDataDecoders*.class",
                "dev/erst/gridgrind/jazzer/support/GeneratedProtocolWorkflow*.class",
                "dev/erst/gridgrind/jazzer/support/HarnessTelemetry*.class",
                "dev/erst/gridgrind/jazzer/support/JazzerGridGrindFuzzData*.class",
                "dev/erst/gridgrind/jazzer/support/OperationSequence*.class",
                "dev/erst/gridgrind/jazzer/support/StyleKindIntrospection*.class",
                "dev/erst/gridgrind/jazzer/support/WorkbookStyleInputs*.class",
                "dev/erst/gridgrind/jazzer/tool/JazzerCli*.class",
                "dev/erst/gridgrind/jazzer/tool/JazzerReportSupport*.class",
                "dev/erst/gridgrind/jazzer/tool/RunMetrics*.class",
            )
    }

    abstract class JazzerToolArgumentsProvider
        @Inject
        constructor() : CommandLineArgumentProvider {
            @get:Input abstract val operation: Property<String>
            @get:Input abstract val projectDirectoryPath: Property<String>
            @get:Optional @get:Input abstract val target: Property<String>
            @get:Optional @get:Input abstract val inputPath: Property<String>
            @get:Optional @get:Input abstract val promotedName: Property<String>
            @get:Optional @get:Input abstract val summarizedTaskName: Property<String>
            @get:Optional @get:Input abstract val logPath: Property<String>
            @get:Optional @get:Input abstract val historyDirectory: Property<String>
            @get:Optional @get:Input abstract val startedAt: Property<String>
            @get:Optional @get:Input abstract val finishedAt: Property<String>
            @get:Optional @get:Input abstract val exitCode: Property<String>
            @get:Optional @get:Input abstract val corpusBeforeFiles: Property<String>
            @get:Optional @get:Input abstract val corpusBeforeBytes: Property<String>
            @get:Input abstract val jsonOutput: Property<Boolean>

            override fun asArguments(): Iterable<String> =
                buildList {
                    add(operation.get())
                    add("--project-dir")
                    add(projectDirectoryPath.get())
                    when (operation.get()) {
                        "status", "report", "list-findings", "list-corpus", "refresh-promoted-metadata" -> {
                            if (target.isPresent) {
                                add("--target")
                                add(target.get())
                            }
                        }

                        "replay" -> {
                            add("--target")
                            add(target.get())
                            add("--input")
                            add(inputPath.get())
                            if (jsonOutput.get()) {
                                add("--json")
                            }
                        }

                        "promote" -> {
                            add("--target")
                            add(target.get())
                            add("--input")
                            add(inputPath.get())
                            add("--name")
                            add(promotedName.get())
                        }

                        "summarize-run" -> {
                            add("--target")
                            add(target.get())
                            add("--task")
                            add(summarizedTaskName.get())
                            add("--log")
                            add(logPath.get())
                            add("--history")
                            add(historyDirectory.get())
                            add("--started-at")
                            add(startedAt.get())
                            add("--finished-at")
                            add(finishedAt.get())
                            add("--exit-code")
                            add(exitCode.get())
                            add("--corpus-before-files")
                            add(corpusBeforeFiles.get())
                            add("--corpus-before-bytes")
                            add(corpusBeforeBytes.get())
                        }
                    }
                }
        }
}
