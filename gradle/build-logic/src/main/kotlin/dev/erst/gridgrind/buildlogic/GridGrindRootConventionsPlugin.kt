package dev.erst.gridgrind.buildlogic

import com.diffplug.gradle.spotless.SpotlessExtension
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.Task
import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.plugins.JacocoPluginExtension
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

            description = providers.gradleProperty("gridgrindDescription").get()
            repositories.mavenCentral()

            allprojects {
                group = providers.gradleProperty("group").get()
                version = providers.gradleProperty("version").get()
            }

            configure<JacocoPluginExtension> {
                toolVersion = libs.findVersion("jacoco").get().requiredVersion
            }

            configure<SpotlessExtension> {
                format("projectFiles") {
                    target(
                        ".gitattributes",
                        ".gitignore",
                        ".dockerignore",
                        "Dockerfile",
                        "**/*.gradle.kts",
                        "**/*.md",
                        "**/*.yml",
                        "gradle.properties",
                        "gradle/**/*.toml",
                        "examples/**/*.json",
                    )
                    targetExclude("**/build/**", "**/.gradle/**")
                    trimTrailingWhitespace()
                    endWithNewline()
                }
            }

            tasks.named("check") {
                dependsOn("spotlessCheck")
            }

            val jacocoAggregatedReport =
                tasks.register<JacocoReport>("jacocoAggregatedReport") {
                group = "verification"
                description = "Aggregates JaCoCo coverage reports from all modules into a single report."

                executionData.from(
                    provider {
                        coverageSubprojects().map { subproject ->
                            subproject.fileTree(subproject.layout.buildDirectory.dir("jacoco").get().asFile) {
                                include("*.exec")
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
                            ) {
                                exclude("**/module-info.class")
                            }
                        }
                    },
                )

                reports {
                    xml.required.set(true)
                    xml.outputLocation.set(layout.buildDirectory.file("reports/jacoco/aggregated/report.xml"))
                    html.required.set(true)
                    html.outputLocation.set(layout.buildDirectory.dir("reports/jacoco/aggregated/html"))
                }
            }

            val coverage =
                tasks.register("coverage") {
                group = "verification"
                description =
                    "Runs tests, enforces coverage thresholds, and generates per-module and aggregated coverage reports."
                dependsOn(jacocoAggregatedReport)
            }

            tasks.register("parity") {
                group = "verification"
                description = "Runs the dedicated Apache POI XSSF parity verification suite."
                dependsOn(":protocol:parityTest")
            }

            gradle.projectsEvaluated {
                val coverageSubprojects = coverageSubprojects()
                val subprojectTestTasks = taskPathsByType(coverageSubprojects, Test::class.java)
                val subprojectCoverageReports = taskPathsByName(coverageSubprojects, "jacocoTestReport")
                val subprojectCoverageVerification =
                    taskPathsByName(coverageSubprojects, "jacocoTestCoverageVerification")
                val subprojectSpotlessTasks = taskPathsByName(coverageSubprojects, "spotlessJava")

                jacocoAggregatedReport.configure {
                    dependsOn(subprojectTestTasks)
                    mustRunAfter(listOf("spotlessProjectFiles") + subprojectSpotlessTasks + subprojectCoverageReports)
                }

                coverage.configure {
                    dependsOn(subprojectCoverageVerification + subprojectCoverageReports)
                }
            }
        }
    }

    private fun Project.coverageSubprojects(): List<Project> =
        subprojects.filter { subproject ->
            subproject.plugins.hasPlugin("jacoco") &&
                subproject.layout.projectDirectory.dir("src/main/java").asFile.isDirectory
        }

    private fun taskPathsByName(subprojects: List<Project>, taskName: String): List<String> =
        subprojects.mapNotNull { subproject ->
            subproject.tasks.findByName(taskName)?.path
        }

    private fun taskPathsByType(subprojects: List<Project>, taskType: Class<out Task>): List<String> =
        subprojects.flatMap { subproject ->
            subproject.tasks.withType(taskType).map(Task::getPath)
        }
}
