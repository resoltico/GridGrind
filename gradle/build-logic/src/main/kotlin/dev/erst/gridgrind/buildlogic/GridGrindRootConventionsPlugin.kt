package dev.erst.gridgrind.buildlogic

import com.diffplug.gradle.spotless.SpotlessExtension
import com.diffplug.spotless.LineEnding
import java.io.File
import org.gradle.api.GradleException
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.Task
import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.plugins.JacocoPluginExtension
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoReport
import org.gradle.testing.jacoco.tasks.JacocoReportsContainer
import org.gradle.kotlin.dsl.configure
import org.gradle.kotlin.dsl.register

class GridGrindRootConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        with(project) {
            pluginManager.apply("base")
            pluginManager.apply("jacoco")
            pluginManager.apply("com.diffplug.spotless")

            val libs = versionCatalog()

            description = providers.gradleProperty("gridgrindDescription").get()
            configureGridGrindRepositories()

            allprojects {
                group = providers.gradleProperty("group").get()
                version = providers.gradleProperty("version").get()
            }

            configure<JacocoPluginExtension> {
                toolVersion = libs.findVersion("jacoco").get().requiredVersion
            }

            configure<SpotlessExtension> {
                lineEndings = LineEnding.UNIX
                format("projectFiles") { formatExtension ->
                    formatExtension.target(*projectFileTargets().toTypedArray())
                    formatExtension.trimTrailingWhitespace()
                    formatExtension.endWithNewline()
                }
            }

            val verifyExplicitImports =
                tasks.register("verifyExplicitImports") { verifyTask ->
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

            tasks.named("check") { checkTask ->
                checkTask.dependsOn("spotlessCheck")
                checkTask.dependsOn(verifyExplicitImports)
            }

            val jacocoAggregatedReport =
                tasks.register<JacocoReport>("jacocoAggregatedReport") {
                    group = "verification"
                    description = "Aggregates JaCoCo coverage reports from all modules into a single report."

                    executionData.from(
                        provider {
                            coverageSubprojects().flatMap { subproject ->
                                subproject.tasks.withType(Test::class.java).map { testTask ->
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
                                ) { patternFilterable ->
                                    patternFilterable.exclude("**/module-info.class")
                                }
                            }
                        },
                    )

                    reports { reports: JacocoReportsContainer ->
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
                tasks.register("coverage") { coverageTask ->
                    coverageTask.group = "verification"
                    coverageTask.description =
                        "Runs tests, enforces coverage thresholds, and generates per-module and aggregated coverage reports."
                    coverageTask.dependsOn(jacocoAggregatedReport)
                }

            tasks.register("parity") { parityTask ->
                parityTask.group = "verification"
                parityTask.description = "Runs the dedicated Apache POI XSSF parity verification suite."
                parityTask.dependsOn(":executor:parityTest")
            }

            gradle.projectsEvaluated {
                val coverageSubprojects = coverageSubprojects()
                val subprojectTestTasks = taskPathsByType(coverageSubprojects, Test::class.java)
                val subprojectCoverageReports = taskPathsByName(coverageSubprojects, "jacocoTestReport")
                val subprojectCoverageVerification =
                    taskPathsByName(coverageSubprojects, "jacocoTestCoverageVerification")
                val subprojectSpotlessTasks = taskPathsByName(coverageSubprojects, "spotlessJava")
                val projectFileSpotlessTasks =
                    listOfNotNull(
                        tasks.findByName("spotlessProjectFiles"),
                        tasks.findByName("spotlessProjectFilesCheck"),
                    )
                val projectFileSpotlessTaskPaths = projectFileSpotlessTasks.map(Task::getPath).toSet()
                val allRootAndSubprojectTasks =
                    (listOf(this@with) + subprojects).flatMap { candidateProject ->
                        candidateProject.tasks.toList()
                    }

                jacocoAggregatedReport.configure { report ->
                    report.dependsOn(subprojectTestTasks)
                    report.mustRunAfter(
                        listOf("spotlessProjectFiles") + subprojectSpotlessTasks + subprojectCoverageReports,
                    )
                }

                coverage.configure { coverageTask ->
                    coverageTask.dependsOn(subprojectCoverageVerification + subprojectCoverageReports)
                }

                // Keep repository-wide project-file formatting on a stable file tree. The target set spans
                // the whole checkout, so letting compile/test tasks create build outputs in parallel can make
                // Spotless walk paths that appear or disappear mid-scan.
                projectFileSpotlessTasks.forEach { projectFileTask ->
                    allRootAndSubprojectTasks
                        .asSequence()
                        .filter { candidateTask -> candidateTask.path !in projectFileSpotlessTaskPaths }
                        .forEach { candidateTask -> candidateTask.mustRunAfter(projectFileTask) }
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
            add(projectFileTree(".codex") { include("**/*.md") })
            add(projectFileTree("docs") { include("**/*.md") })
            add(projectFileTree(".github") { include("**/*.yml") })
            add(projectFileTree("examples") { include("**/*.json") })
        }

    private fun Project.repositoryProjectFiles(): List<File> =
        buildList {
            add(rootFile(".gitattributes"))
            add(rootFile(".gitignore"))
            add(rootFile(".dockerignore"))
            add(rootFile("AGENTS.md"))
            add(rootFile("CHANGELOG.md"))
            add(rootFile("Dockerfile"))
            add(rootFile("LICENSE"))
            add(rootFile("LICENSE-APACHE-2.0"))
            add(rootFile("LICENSE-BSD-3-CLAUSE"))
            add(rootFile("NOTICE"))
            add(rootFile("PATENTS.md"))
            add(rootFile("README.md"))
            add(rootFile("build.gradle.kts"))
            add(rootFile("check.sh"))
            add(rootFile("gradle.properties"))
            add(rootFile("settings.gradle.kts"))
            add(rootFile("authoring-java/build.gradle.kts"))
            add(rootFile("cli/build.gradle.kts"))
            add(rootFile("contract/build.gradle.kts"))
            add(rootFile("engine/build.gradle.kts"))
            add(rootFile("excel-foundation/build.gradle.kts"))
            add(rootFile("executor/build.gradle.kts"))
            add(rootFile("gradle/build-logic/build.gradle.kts"))
            add(rootFile("gradle/build-logic/settings.gradle.kts"))
            add(rootFile("gradle/libs.versions.toml"))
            add(rootFile("jazzer/README.md"))
            add(rootFile("jazzer/build.gradle.kts"))
            add(rootFile("jazzer/settings.gradle.kts"))
        }.filter(File::isFile)

    private fun Project.projectFileTree(
        path: String,
        configureTree: org.gradle.api.file.ConfigurableFileTree.() -> Unit,
    ) = fileTree(layout.projectDirectory.dir(path)) { fileTree ->
        fileTree.configureTree()
    }

    private fun Project.rootFile(path: String): File =
        layout.projectDirectory.file(path).asFile

    private fun taskPathsByName(subprojects: List<Project>, taskName: String): List<String> =
        subprojects.mapNotNull { subproject ->
            subproject.tasks.findByName(taskName)?.path
        }

    private fun taskPathsByType(subprojects: List<Project>, taskType: Class<out Task>): List<String> =
        subprojects.flatMap { subproject ->
            subproject.tasks.withType(taskType).map(Task::getPath)
        }

    companion object {
        private val WILDCARD_IMPORT_PATTERN = Regex("""^import(?:\s+static)?\s+[\w.]+\.\*;$""")
    }
}
