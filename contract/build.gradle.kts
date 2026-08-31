import org.gradle.api.tasks.testing.Test
import org.gradle.api.tasks.JavaExec
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport

plugins {
    `java-library`
    id("gridgrind.java-conventions")
    id("gridgrind.mutation-conventions")
}

description = "Canonical GridGrind contract model, metadata registry, and JSON codecs"

gridgrindMutation {
    targetClasses.set(
        setOf(
            "dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation*",
            "dev.erst.gridgrind.contract.dto.ProtocolRgbColorSupport*",
            "dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput*",
            "dev.erst.gridgrind.contract.dto.DataValidationRuleInput*",
            "dev.erst.gridgrind.contract.catalog.FieldConstraint*",
            "dev.erst.gridgrind.contract.catalog.CatalogStepTemplateDefaults*",
            "dev.erst.gridgrind.contract.catalog.CatalogStepTemplateSupport*",
            "dev.erst.gridgrind.contract.json.RequestUtf8DecodeResult*",
            "dev.erst.gridgrind.contract.step.WorkbookStaticMaterializationValidation*",
        ),
    )
    targetTests.set(
        setOf(
            "dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidationTest",
            "dev.erst.gridgrind.contract.dto.AdvancedMutationProtocolTypesTest",
            "dev.erst.gridgrind.contract.dto.ConditionalFormattingDataBarBoundsTest",
            "dev.erst.gridgrind.contract.dto.ProtocolDefaultingCoverageTest",
            "dev.erst.gridgrind.contract.dto.ProtocolResidualCoverageTest",
            "dev.erst.gridgrind.contract.catalog.FieldConstraintTest",
            "dev.erst.gridgrind.contract.catalog.CatalogStepTemplateSupportTest",
            "dev.erst.gridgrind.contract.json.RequestSyntaxSupportTest",
            "dev.erst.gridgrind.contract.dto.ProtocolRgbColorSupportTest",
            "dev.erst.gridgrind.contract.step.WorkbookStaticMaterializationValidationTest",
            "dev.erst.gridgrind.contract.step.WorkbookStaticRequestContractTest",
        ),
    )
    mutationThreshold.set(100)
    coverageThreshold.set(100)
    testStrengthThreshold.set(100)
    maxSurviving.set(0)
}

val downstreamCoverageProjects =
    listOf(project(":engine"), project(":executor"), project(":authoring-java"), project(":cli"))

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
    api(project(":excel-foundation"))
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

tasks.withType<Test>().configureEach {
    systemProperty("gridgrind.root.dir", rootProject.projectDir.absolutePath)
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
