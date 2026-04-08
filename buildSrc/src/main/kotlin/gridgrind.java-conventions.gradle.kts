/**
 * Convention plugin applied to every GridGrind Java subproject.
 *
 * Centralises:
 *  - Java toolchain (version read once from gradle.properties)
 *  - sources JAR production
 *  - JUnit Platform test runner
 *  - JaCoCo report + coverage-verification tasks (100 % line, 100 % branch)
 */
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.tasks.bundling.Jar
import org.gradle.api.tasks.testing.Test
import org.gradle.api.tasks.testing.TestDescriptor
import org.gradle.api.tasks.testing.TestListener
import org.gradle.api.tasks.testing.TestResult
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport

plugins {
    `java-base`
    jacoco
}

val gridgrindJavaVersion: Int =
    providers.gradleProperty("gridgrindJavaVersion").map(String::toInt).get()

extensions.configure<JavaPluginExtension> {
    toolchain {
        languageVersion = JavaLanguageVersion.of(gridgrindJavaVersion)
    }
    modularity.inferModulePath = true
    withSourcesJar()
}

tasks.withType<Jar>().configureEach {
    manifest {
        attributes(
            "Implementation-Title" to project.name,
            "Implementation-Version" to project.version,
            "Implementation-Vendor" to "Ervins Strauhmanis",
            "Implementation-License" to "MIT",
        )
    }
}

tasks.withType<Test>().configureEach {
    useJUnitPlatform()

    val progressPulseEnabled =
        providers.environmentVariable("GRIDGRIND_TEST_PULSE").map { it == "1" }.orElse(false).get()
    if (progressPulseEnabled) {
        val progressPulseIntervalMillis =
            providers.environmentVariable("GRIDGRIND_TEST_PULSE_INTERVAL_MS")
                .map(String::toLong)
                .orElse(15_000L)
                .get()
        val pulseTaskPath = path
        val pulseProjectPath = project.path
        fun pulseValue(raw: String?): String = raw?.replace(Regex("\\s+"), "_") ?: "unknown"

        addTestListener(
            object : TestListener {
                private var completedClasses = 0L
                private var completedTests = 0L
                private var failedTests = 0L
                private var skippedTests = 0L
                private var lastPulseAtMillis = 0L

                override fun beforeSuite(suite: TestDescriptor) {
                    val now = System.currentTimeMillis()
                    if (suite.parent == null) {
                        logger.lifecycle(
                            "[GRADLE-TEST-PULSE] task={} project={} phase=start",
                            pulseTaskPath,
                            pulseProjectPath,
                        )
                        lastPulseAtMillis = now
                        return
                    }
                    if (suite.className != null && suite.parent?.className == null) {
                        logger.lifecycle(
                            "[GRADLE-TEST-PULSE] task={} project={} phase=class-start class={}",
                            pulseTaskPath,
                            pulseProjectPath,
                            pulseValue(suite.className),
                        )
                        lastPulseAtMillis = now
                    }
                }

                override fun afterSuite(suite: TestDescriptor, result: TestResult) {
                    val now = System.currentTimeMillis()
                    if (suite.parent == null) {
                        logger.lifecycle(
                            "[GRADLE-TEST-PULSE] task={} project={} phase=finish completedClasses={} completedTests={} failedTests={} skippedTests={} result={}",
                            pulseTaskPath,
                            pulseProjectPath,
                            completedClasses,
                            completedTests,
                            failedTests,
                            skippedTests,
                            result.resultType,
                        )
                        lastPulseAtMillis = now
                        return
                    }
                    if (suite.className != null && suite.parent?.className == null) {
                        completedClasses += 1
                        logger.lifecycle(
                            "[GRADLE-TEST-PULSE] task={} project={} phase=class-complete completedClasses={} completedTests={} failedTests={} skippedTests={} class={} classTests={} classFailedTests={} classSkippedTests={} result={}",
                            pulseTaskPath,
                            pulseProjectPath,
                            completedClasses,
                            completedTests,
                            failedTests,
                            skippedTests,
                            pulseValue(suite.className),
                            result.testCount,
                            result.failedTestCount,
                            result.skippedTestCount,
                            result.resultType,
                        )
                        lastPulseAtMillis = now
                    }
                }

                override fun beforeTest(testDescriptor: TestDescriptor) = Unit

                override fun afterTest(testDescriptor: TestDescriptor, result: TestResult) {
                    completedTests += 1
                    when (result.resultType) {
                        TestResult.ResultType.FAILURE -> failedTests += 1
                        TestResult.ResultType.SKIPPED -> skippedTests += 1
                        TestResult.ResultType.SUCCESS -> Unit
                    }

                    val now = System.currentTimeMillis()
                    if (now - lastPulseAtMillis >= progressPulseIntervalMillis) {
                        logger.lifecycle(
                            "[GRADLE-TEST-PULSE] task={} project={} phase=test-progress completedClasses={} completedTests={} failedTests={} skippedTests={} class={} test={} result={}",
                            pulseTaskPath,
                            pulseProjectPath,
                            completedClasses,
                            completedTests,
                            failedTests,
                            skippedTests,
                            pulseValue(testDescriptor.className),
                            pulseValue(testDescriptor.name),
                            result.resultType,
                        )
                        lastPulseAtMillis = now
                    }
                }
            },
        )
    }
}

tasks.named<JacocoReport>("jacocoTestReport") {
    dependsOn(tasks.withType<Test>())
    reports {
        xml.required = true
        html.required = true
    }
}

tasks.named<JacocoCoverageVerification>("jacocoTestCoverageVerification") {
    dependsOn(tasks.withType<Test>())
    violationRules {
        rule {
            limit {
                counter = "LINE"
                value = "COVEREDRATIO"
                minimum = "1.0".toBigDecimal()
            }
            limit {
                counter = "BRANCH"
                value = "COVEREDRATIO"
                minimum = "1.0".toBigDecimal()
            }
        }
    }
}
