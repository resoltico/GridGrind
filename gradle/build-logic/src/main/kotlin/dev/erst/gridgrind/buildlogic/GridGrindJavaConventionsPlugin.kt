package dev.erst.gridgrind.buildlogic

import com.diffplug.gradle.spotless.SpotlessExtension
import java.math.BigDecimal
import org.gradle.api.Action
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.artifacts.VersionCatalogsExtension
import org.gradle.api.plugins.quality.Pmd
import org.gradle.api.plugins.quality.PmdExtension
import org.gradle.api.tasks.compile.JavaCompile
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.tasks.bundling.Jar
import org.gradle.api.tasks.testing.Test
import org.gradle.jvm.toolchain.JavaLanguageVersion
import org.gradle.kotlin.dsl.configure
import org.gradle.kotlin.dsl.getByType
import org.gradle.testing.jacoco.plugins.JacocoPluginExtension
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport
import org.gradle.testing.jacoco.tasks.rules.JacocoLimit
import org.gradle.testing.jacoco.tasks.rules.JacocoViolationRule
import org.gradle.testing.jacoco.tasks.rules.JacocoViolationRulesContainer
import net.ltgt.gradle.errorprone.errorprone

class GridGrindJavaConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        with(project) {
            pluginManager.apply("java-base")
            pluginManager.apply("jacoco")
            pluginManager.apply("com.diffplug.spotless")
            pluginManager.apply("net.ltgt.errorprone")
            pluginManager.apply("pmd")

            repositories.mavenCentral()

            val libs = extensions.getByType<VersionCatalogsExtension>().named("libs")
            val repositoryLayout = GridGrindRepositoryLayout.locate(this)
            val gridgrindJavaVersion =
                providers.gradleProperty("gridgrindJavaVersion").map(String::toInt).get()

            pluginManager.withPlugin("java") {
                val javaExtension = extensions.getByType(JavaPluginExtension::class.java)
                javaExtension.toolchain.languageVersion.set(JavaLanguageVersion.of(gridgrindJavaVersion))
                javaExtension.modularity.inferModulePath.set(true)
                javaExtension.withSourcesJar()
            }

            dependencies.add("errorprone", libs.findLibrary("errorprone-core").get().get())

            extensions.configure<SpotlessExtension> {
                java {
                    target("src/*/java/**/*.java")
                    googleJavaFormat(libs.findVersion("google-java-format").get().requiredVersion)
                    removeUnusedImports()
                    formatAnnotations()
                }
            }

            extensions.configure<PmdExtension> {
                toolVersion = libs.findVersion("pmd").get().requiredVersion
                isConsoleOutput = true
                isIgnoreFailures = false
                rulesMinimumPriority.set(3)
                setRuleSetFiles(files(repositoryLayout.productionPmdRuleset))
                setRuleSets(emptyList())
            }

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

            tasks.withType(JavaCompile::class.java).configureEach(
                object : Action<JavaCompile> {
                    override fun execute(javaCompile: JavaCompile) {
                        javaCompile.options.errorprone.disableWarningsInGeneratedCode.set(true)
                        javaCompile.options.errorprone.error(
                            "BadImport",
                            "BoxedPrimitiveConstructor",
                            "CheckReturnValue",
                            "EqualsIncompatibleType",
                            "JavaLangClash",
                            "MissingCasesInEnumSwitch",
                            "MissingOverride",
                            "ReferenceEquality",
                            "StringCaseLocaleUsage",
                        )
                    }
                },
            )

            tasks.withType(Pmd::class.java).configureEach(
                object : Action<Pmd> {
                    override fun execute(pmd: Pmd) {
                        pmd.reports.xml.required.set(true)
                        pmd.reports.html.required.set(true)
                    }
                },
            )

            tasks.withType(Pmd::class.java)
                .matching { task -> task.name == "pmdTest" }
                .configureEach(
                    object : Action<Pmd> {
                        override fun execute(pmd: Pmd) {
                            pmd.ruleSetFiles = files(repositoryLayout.testPmdRuleset)
                            pmd.ruleSets = emptyList()
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

            extensions.configure<JacocoPluginExtension> {
                toolVersion = libs.findVersion("jacoco").get().requiredVersion
            }

            tasks.named("jacocoTestReport", JacocoReport::class.java).configure(
                object : Action<JacocoReport> {
                    override fun execute(report: JacocoReport) {
                        report.dependsOn(tasks.withType(Test::class.java))
                        report.executionData.from(
                            tasks.withType(Test::class.java).map { testTask ->
                                testTask.extensions.getByType(JacocoTaskExtension::class.java).destinationFile
                            },
                        )
                        report.reports.xml.required.set(true)
                        report.reports.html.required.set(true)
                    }
                },
            )

            tasks.named("check").configure {
                dependsOn("spotlessCheck")
                dependsOn("jacocoTestCoverageVerification")
            }

            tasks.named("jacocoTestCoverageVerification", JacocoCoverageVerification::class.java).configure(
                object : Action<JacocoCoverageVerification> {
                    override fun execute(verification: JacocoCoverageVerification) {
                        verification.dependsOn(tasks.withType(Test::class.java))
                        verification.executionData.from(
                            tasks.withType(Test::class.java).map { testTask ->
                                testTask.extensions.getByType(JacocoTaskExtension::class.java).destinationFile
                            },
                        )
                        verification.violationRules(
                            object : Action<JacocoViolationRulesContainer> {
                                override fun execute(rules: JacocoViolationRulesContainer) {
                                    rules.rule(
                                        object : Action<JacocoViolationRule> {
                                            override fun execute(rule: JacocoViolationRule) {
                                                rule.limit(
                                                    object : Action<JacocoLimit> {
                                                        override fun execute(limit: JacocoLimit) {
                                                            limit.counter = "LINE"
                                                            limit.value = "COVEREDRATIO"
                                                            limit.minimum = BigDecimal("1.0")
                                                        }
                                                    },
                                                )
                                                rule.limit(
                                                    object : Action<JacocoLimit> {
                                                        override fun execute(limit: JacocoLimit) {
                                                            limit.counter = "BRANCH"
                                                            limit.value = "COVEREDRATIO"
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
