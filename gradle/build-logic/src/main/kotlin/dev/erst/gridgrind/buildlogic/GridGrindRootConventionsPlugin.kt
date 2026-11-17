package dev.erst.gridgrind.buildlogic

import com.diffplug.gradle.spotless.SpotlessExtension
import com.diffplug.spotless.LineEnding
import java.io.File
import org.gradle.api.GradleException
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.plugins.JacocoPluginExtension
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoReport
import org.gradle.kotlin.dsl.configure
import org.gradle.kotlin.dsl.register

class GridGrindRootConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        with(project) {
            pluginManager.apply("base")
            pluginManager.apply("jacoco")
            pluginManager.apply("com.diffplug.spotless")

            val libs = versionCatalog()
            val repositoryLayout = GridGrindRepositoryLayout.locate(this)

            description = providers.gradleProperty("gridgrindDescription").get()
            configureGridGrindRepositories()

            allprojects { val candidate = this
                candidate.dependencyLocking { lockAllConfigurations() }
                candidate.group = providers.gradleProperty("group").get()
                candidate.version = providers.gradleProperty("version").get()
            }

            configure<JacocoPluginExtension> {
                toolVersion = libs.findVersion("jacoco").get().requiredVersion
            }

            configure<SpotlessExtension> {
                lineEndings = LineEnding.UNIX
                format("projectFiles") { val formatExtension = this
                    formatExtension.target(*projectFileTargets().toTypedArray())
                    formatExtension.trimTrailingWhitespace()
                    formatExtension.endWithNewline()
                }
            }

            val verifyExplicitImports =
                tasks.register("verifyExplicitImports") { val verifyTask = this
                    verifyTask.group = "verification"
                    verifyTask.description =
                        "Fails when handwritten production Java/Kotlin sources use wildcard imports."

                    val rootDirectory = layout.projectDirectory.asFile
                    val sourceRoots = explicitImportSourceRoots()
                    verifyTask.inputs.files(sourceRoots)

                    verifyTask.doLast {
                        val violations =
                            buildList {
                                sourceRoots.forEach { sourceRoot ->
                                    sourceRoot.walkTopDown()
                                        .filter { file ->
                                            file.isFile &&
                                                (file.extension == "java" || file.extension == "kt")
                                        }
                                        .forEach { file ->
                                            file.readLines().forEachIndexed { index, line ->
                                                val trimmed = line.trim()
                                                if (WILDCARD_IMPORT_PATTERN.matches(trimmed)) {
                                                    add(
                                                        "${file.relativeTo(rootDirectory).invariantSeparatorsPath}:${index + 1}: $trimmed",
                                                    )
                                                }
                                            }
                                        }
                                }
                            }

                        if (violations.isNotEmpty()) {
                            throw GradleException(
                                buildString {
                                    appendLine(
                                        "Wildcard imports are forbidden in handwritten production sources.",
                                    )
                                    appendLine("Replace these imports with explicit symbols:")
                                    violations.forEach { violation -> appendLine(" - $violation") }
                                },
                            )
                        }
                    }
                }

            val verifyNoLegacyBuildSrc =
                tasks.register("verifyNoLegacyBuildSrc") { val verifyTask = this
                    verifyTask.group = "verification"
                    verifyTask.description =
                        "Fails when buildSrc is present in an owned Gradle build or project."
                    val ownedBuildSrcDirectories =
                        (listOf(repositoryLayout.repositoryRoot) +
                            subprojects.map { it.projectDir } +
                            listOf(repositoryLayout.repositoryRoot.resolve("gradle/build-logic")))
                            .map { it.resolve("buildSrc") }
                    verifyTask.inputs.files(ownedBuildSrcDirectories)
                    verifyTask.doLast {
                        val legacyBuildSrcDirectories = ownedBuildSrcDirectories.filter(File::isDirectory)


                        if (legacyBuildSrcDirectories.isNotEmpty()) {
                            throw GradleException(
                                buildString {
                                    appendLine(
                                        "Legacy buildSrc directories are forbidden; GridGrind uses the shared included build under gradle/build-logic.",
                                    )
                                    appendLine("Remove these directories instead of reactivating implicit build logic:")
                                    legacyBuildSrcDirectories.forEach { directory ->
                                        appendLine(
                                            " - ${directory.relativeTo(repositoryLayout.repositoryRoot).invariantSeparatorsPath}",
                                        )
                                    }
                                },
                            )
                        }
                    }
                }

            tasks.register<ReleaseNotesTask>("prepareReleaseNotes") {
                group = "distribution"
                description = "Extracts the exact versioned CHANGELOG section; never publishes."
                releaseVersion.set(providers.gradleProperty("version"))
                changelog.set(layout.projectDirectory.file("CHANGELOG.md"))
                notes.set(layout.buildDirectory.file("release-notes.md"))
            }

            val verifyBuildLogicTests = gradle.includedBuild("build-logic").task(":test")
            val architectureCheck = registerGridGrindArchitectureCheck()
            registerGridGrindMutationReport()

            val reportRetention = registerGridGrindReportRetention()
            tasks.named("check") { val checkTask = this
                checkTask.finalizedBy(reportRetention)
                checkTask.dependsOn("spotlessCheck")
                checkTask.dependsOn(verifyExplicitImports)
                checkTask.dependsOn(verifyNoLegacyBuildSrc)
                checkTask.dependsOn(verifyBuildLogicTests)
                checkTask.dependsOn(architectureCheck)
            }

            val jacocoAggregatedReport =
                tasks.register<JacocoReport>("jacocoAggregatedReport") {
                    group = "verification"
                    description = "Aggregates JaCoCo coverage reports from all modules into a single report."

                    executionData.from(
                        provider {
                            coverageSubprojects().flatMap { subproject ->
                                subproject.coverageTestTasks().map { testTask ->
                                    testTask.extensions.getByType(JacocoTaskExtension::class.java).destinationFile
                                }
                            }
                        },
                    )
                    sourceDirectories.from(
                        provider {
                            coverageSubprojects().map { subproject ->
                                subproject.layout.projectDirectory.dir("src/main/java").asFile
                            }
                        },
                    )
                    classDirectories.from(
                        provider {
                            coverageSubprojects().map { subproject ->
                                subproject.fileTree(
                                    subproject.layout.buildDirectory.dir("classes/java/main").get().asFile,
                                ) { val patternFilterable = this
                                    patternFilterable.exclude("**/module-info.class")
                                }
                            }
                        },
                    )

                    reports { val reports = this
                        reports.xml.required.set(true)
                        reports.xml.outputLocation.set(
                            layout.buildDirectory.file("reports/jacoco/aggregated/report.xml"),
                        )
                        reports.html.required.set(true)
                        reports.html.outputLocation.set(
                            layout.buildDirectory.dir("reports/jacoco/aggregated/html"),
                        )
                    }
                }

            val coverage =
                tasks.register("coverage") { val coverageTask = this
                    coverageTask.group = "verification"
                    coverageTask.description =
                        "Runs tests and generates per-module and aggregated coverage reports."
                    coverageTask.dependsOn(jacocoAggregatedReport)
                }

            tasks.register("parity") { val parityTask = this
                parityTask.group = "verification"
                parityTask.description = "Runs the dedicated Apache POI XSSF parity verification suite."
                parityTask.dependsOn(":engine:parityTest")
            }

            gradle.projectsEvaluated {
                val coverageSubprojects = coverageSubprojects()
                val subprojectCoverageReports = taskPathsByName(coverageSubprojects, "jacocoTestReport")
                jacocoAggregatedReport.configure { val report = this
                    report.dependsOn(subprojectCoverageReports)

                }

                coverage.configure { val coverageTask = this
                    coverageTask.dependsOn(subprojectCoverageReports)
                }


            }
        }
    }

    private fun Project.coverageSubprojects(): List<Project> =
        subprojects.filter { subproject ->
            subproject.plugins.hasPlugin("jacoco") &&
                subproject.layout.projectDirectory.dir("src/main/java").asFile.isDirectory
        }

    private fun Project.explicitImportSourceRoots(): List<File> =
        buildList {
            add(layout.projectDirectory.dir("src/main/java").asFile)
            add(layout.projectDirectory.dir("src/main/kotlin").asFile)
            subprojects.forEach { subproject ->
                add(subproject.layout.projectDirectory.dir("src/main/java").asFile)
                add(subproject.layout.projectDirectory.dir("src/main/kotlin").asFile)
            }
            add(layout.projectDirectory.dir("gradle/build-logic/src/main/java").asFile)
            add(layout.projectDirectory.dir("gradle/build-logic/src/main/kotlin").asFile)
        }.distinct().filter(File::isDirectory)

    private fun Project.projectFileTargets(): List<Any> =
        buildList {
            repositoryProjectFiles().forEach(::add)
            add(projectFileTree("docs") { include("**/*.md") })
            add(projectFileTree(".github") { include("**/*.yml") })
            add(projectFileTree("examples") { include("**/*.json") })
        }

    private fun Project.repositoryProjectFiles(): List<File> =
        buildList {
            add(rootFile(".gitattributes"))
            add(rootFile(".gitignore"))
            add(rootFile(".dockerignore"))
            add(rootFile("CHANGELOG.md"))
            add(rootFile("Dockerfile"))
            add(rootFile("LICENSE"))
            add(rootFile("PROJECT-NOTICE.md"))
            add(rootFile("LICENSE-COMMERCIAL.md"))
            add(rootFile("CONTRIBUTING.md"))
            add(rootFile(".github/pull_request_template.md"))
            add(rootFile("LICENSE-APACHE-2.0"))
            add(rootFile("LICENSE-BSD-2-CLAUSE"))
            add(rootFile("LICENSE-BSD-3-CLAUSE"))
            add(rootFile("NOTICE"))
            add(rootFile("PATENTS.md"))
            add(rootFile("README.md"))
            add(rootFile("build.gradle.kts"))
            add(rootFile("check.sh"))
            add(rootFile("gradle.properties"))
            add(rootFile("settings.gradle.kts"))
            add(rootFile("core/build.gradle.kts"))
            add(rootFile("json/build.gradle.kts"))
            add(rootFile("cli/build.gradle.kts"))
            add(rootFile("engine/build.gradle.kts"))
            add(rootFile("architecture-tests/build.gradle.kts"))
            add(rootFile("gradle/build-logic/build.gradle.kts"))
            add(rootFile("gradle/build-logic/settings.gradle.kts"))
            add(rootFile("gradle/libs.versions.toml"))
        }.filter(File::isFile)

    private fun Project.projectFileTree(
        path: String,
        configureTree: org.gradle.api.file.ConfigurableFileTree.() -> Unit,
    ) = fileTree(layout.projectDirectory.dir(path)) { val fileTree = this
        fileTree.configureTree()
    }

    private fun Project.rootFile(path: String): File =
        layout.projectDirectory.file(path).asFile

    private fun taskPathsByName(subprojects: List<Project>, taskName: String): List<String> =
        subprojects.mapNotNull { subproject ->
            subproject.tasks.findByName(taskName)?.path
        }

    companion object {
        private val WILDCARD_IMPORT_PATTERN = Regex("""^import(?:\s+static)?\s+[\w.]+\.\*;$""")
    }
}
