import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport

plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "Canonical GridGrind contract model, metadata registry, and JSON codecs"

val downstreamCoverageProjects = listOf(project(":executor"), project(":authoring-java"), project(":cli"))

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

dependencies {
    implementation(project(":engine"))
    // Jackson 3.x still owns annotations via the Jackson 2.x coordinates and package namespace.
    api(libs.jackson.annotations)
    implementation(libs.jackson.databind)
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testImplementation(libs.poi.ooxml.full)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}

tasks.named<JacocoReport>("jacocoTestReport") {
    dependsOn(downstreamCoverageTaskPaths())
    executionData.from(downstreamCoverageExecutionData())
}

tasks.named<JacocoCoverageVerification>("jacocoTestCoverageVerification") {
    dependsOn(downstreamCoverageTaskPaths())
    executionData.from(downstreamCoverageExecutionData())
}
