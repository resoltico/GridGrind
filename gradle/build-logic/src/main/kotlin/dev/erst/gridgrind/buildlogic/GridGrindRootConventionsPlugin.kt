package dev.erst.gridgrind.buildlogic

import com.diffplug.gradle.spotless.SpotlessExtension
import org.gradle.api.Plugin
import org.gradle.api.Project
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

            tasks.register<JacocoReport>("jacocoAggregatedReport") {
                group = "verification"
                description = "Aggregates JaCoCo coverage reports from all modules into a single report."

                dependsOn(":engine:test", ":protocol:test", ":cli:test")
                mustRunAfter(
                    "spotlessProjectFiles",
                    ":engine:spotlessJava",
                    ":protocol:spotlessJava",
                    ":cli:spotlessJava",
                    ":engine:jacocoTestReport",
                    ":protocol:jacocoTestReport",
                    ":cli:jacocoTestReport",
                )

                executionData.from(
                    subprojects.map { subproject -> file("${subproject.projectDir}/build/jacoco/test.exec") },
                )
                sourceDirectories.from(
                    subprojects.map { subproject -> file("${subproject.projectDir}/src/main/java") },
                )
                classDirectories.from(
                    subprojects.map { subproject ->
                        fileTree("${subproject.projectDir}/build/classes/java/main") {
                            exclude("**/module-info.class")
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

            tasks.register("coverage") {
                group = "verification"
                description =
                    "Runs tests, enforces coverage thresholds, and generates per-module and aggregated coverage reports."
                dependsOn(
                    ":engine:jacocoTestCoverageVerification",
                    ":protocol:jacocoTestCoverageVerification",
                    ":cli:jacocoTestCoverageVerification",
                    ":engine:jacocoTestReport",
                    ":protocol:jacocoTestReport",
                    ":cli:jacocoTestReport",
                    "jacocoAggregatedReport",
                )
            }

            tasks.register("parity") {
                group = "verification"
                description = "Runs the dedicated Apache POI XSSF parity verification suite."
                dependsOn(":protocol:parityTest")
            }
        }
    }
}
