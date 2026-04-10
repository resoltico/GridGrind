package dev.erst.gridgrind.buildlogic

import java.math.BigDecimal
import org.gradle.api.Action
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.tasks.bundling.Jar
import org.gradle.api.tasks.testing.Test
import org.gradle.jvm.toolchain.JavaLanguageVersion
import org.gradle.testing.jacoco.tasks.rules.JacocoViolationRule
import org.gradle.testing.jacoco.tasks.rules.JacocoViolationRulesContainer
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport
import org.gradle.testing.jacoco.tasks.rules.JacocoLimit

class GridGrindJavaConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        with(project) {
            pluginManager.apply("java-base")
            pluginManager.apply("jacoco")

            val gridgrindJavaVersion =
                providers.gradleProperty("gridgrindJavaVersion").map(String::toInt).get()

            val javaExtension = extensions.getByType(JavaPluginExtension::class.java)
            javaExtension.toolchain.languageVersion.set(JavaLanguageVersion.of(gridgrindJavaVersion))
            javaExtension.modularity.inferModulePath.set(true)
            javaExtension.withSourcesJar()

            tasks.withType(Jar::class.java).configureEach(
                object : Action<Jar> {
                    override fun execute(jar: Jar) {
                    jar.manifest.attributes(
                        mapOf(
                            "Implementation-Title" to project.name,
                            "Implementation-Version" to project.version,
                            "Implementation-Vendor" to "Ervins Strauhmanis",
                            "Implementation-License" to "MIT",
                        ),
                    )
                    }
                },
            )

            tasks.withType(Test::class.java).configureEach(
                object : Action<Test> {
                    override fun execute(test: Test) {
                        test.useJUnitPlatform()

                        val progressPulseEnabled =
                            providers.environmentVariable("GRIDGRIND_TEST_PULSE")
                                .map { it == "1" }
                                .orElse(false)
                                .get()
                        if (progressPulseEnabled) {
                            val progressPulseIntervalMillis =
                                providers.environmentVariable("GRIDGRIND_TEST_PULSE_INTERVAL_MS")
                                    .map(String::toLong)
                                    .orElse(15_000L)
                                    .get()
                            val pulseTaskPath = test.path
                            val pulseProjectPath = project.path
                            test.doFirst {
                                test.addTestListener(
                                    GradleTestPulseListener(
                                        logger = test.logger,
                                        taskPath = pulseTaskPath,
                                        projectPath = pulseProjectPath,
                                        pulseIntervalMillis = progressPulseIntervalMillis,
                                    ),
                                )
                            }
                        }
                    }
                },
            )

            tasks.named("jacocoTestReport", JacocoReport::class.java).configure(
                object : Action<JacocoReport> {
                    override fun execute(report: JacocoReport) {
                        report.dependsOn(tasks.withType(Test::class.java))
                        report.reports.xml.required.set(true)
                        report.reports.html.required.set(true)
                    }
                },
            )

            tasks.named("jacocoTestCoverageVerification", JacocoCoverageVerification::class.java).configure(
                object : Action<JacocoCoverageVerification> {
                    override fun execute(verification: JacocoCoverageVerification) {
                        verification.dependsOn(tasks.withType(Test::class.java))
                        verification.violationRules(
                            object : Action<JacocoViolationRulesContainer> {
                                override fun execute(rules: JacocoViolationRulesContainer) {
                                    rules.rule(
                                        object : Action<JacocoViolationRule> {
                                            override fun execute(rule: JacocoViolationRule) {
                                                rule.limit(
                                                    object : Action<JacocoLimit> {
                                                        override fun execute(limit: JacocoLimit) {
                                                            limit.minimum = BigDecimal("1.0")
                                                        }
                                                    },
                                                )
                                            }
                                        },
                                    )
                                }
                            },
                        )
                    }
                },
            )
        }
    }
}
