import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport

plugins {
    `java-library`
    `java-test-fixtures`
    id("gridgrind.java-conventions")
    id("gridgrind.mutation-conventions")
}

description = "Core GridGrind workbook automation and request execution engine"

gridgrindMutation {
    targetClasses.set(
        setOf(
            "dev.erst.gridgrind.engine.runtime.FormulaOriginTracker*",
            "dev.erst.gridgrind.engine.runtime.RequestPreflightPaths*",
            "dev.erst.gridgrind.excel.validation.ExcelDataValidationComparisonOperatorPoiBridge*",
        ),
    )
    targetTests.set(
        setOf(
            "dev.erst.gridgrind.engine.runtime.FormulaOriginTrackerTest",
            "dev.erst.gridgrind.engine.runtime.RequestPreflightTest",
            "dev.erst.gridgrind.excel.validation.ExcelDataValidationComparisonOperatorPoiBridgeTest",
        ),
    )
    mutationThreshold.set(100)
    coverageThreshold.set(100)
    testStrengthThreshold.set(100)
    maxSurviving.set(0)
}

tasks.named("pitest") {
    mustRunAfter(":contract:pitest")
}

// Engine behavior is exercised through both runtime adapters; coverage evidence must include both.
val downstreamCoverageProjects = listOf(project(":cli"), project(":executor"))

fun downstreamCoverageTaskPaths(): List<String> =
    downstreamCoverageProjects.flatMap { downstreamProject ->
        downstreamProject.tasks.withType<Test>().map { task -> task.path }
    }

fun downstreamCoverageExecutionData() =
    provider {
        downstreamCoverageProjects.flatMap { downstreamProject ->
            downstreamProject.tasks.withType<Test>().map { testTask ->
                testTask.extensions.getByType(JacocoTaskExtension::class.java).destinationFile
            }
        }
    }

fun localCoverageExecutionData() =
    provider {
        listOf(
            tasks.named<Test>("test")
                .get()
                .extensions
                .getByType(JacocoTaskExtension::class.java)
                .destinationFile
        )
    }

dependencies {
    api(project(":contract"))
    api(project(":excel-foundation"))
    implementation(libs.jakarta.activation)
    implementation(libs.jakarta.xml.bind)
    implementation(libs.poi.ooxml)
    implementation(libs.poi.ooxml.full)
    implementation(libs.xmlsec)
    runtimeOnly(libs.bcpkix)
    runtimeOnly(libs.bcprov)
    runtimeOnly(libs.bcutil)
    testFixturesImplementation(libs.jakarta.activation)
    testFixturesImplementation(libs.jakarta.xml.bind)
    testFixturesImplementation(libs.bcpkix)
    testFixturesImplementation(libs.bcprov)
    testFixturesImplementation(libs.bcutil)
    testFixturesImplementation(libs.poi.ooxml)
    testFixturesImplementation(libs.poi.ooxml.full)
    testFixturesImplementation(libs.xmlsec)
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}

tasks.named<JacocoReport>("jacocoTestReport") {
    dependsOn(tasks.named<Test>("test"))
    executionData.from(localCoverageExecutionData())
    dependsOn(downstreamCoverageTaskPaths())
    executionData.from(downstreamCoverageExecutionData())
}

tasks.named<JacocoCoverageVerification>("jacocoTestCoverageVerification") {
    dependsOn(tasks.named<Test>("test"))
    executionData.from(localCoverageExecutionData())
    dependsOn(downstreamCoverageTaskPaths())
    executionData.from(downstreamCoverageExecutionData())
}
