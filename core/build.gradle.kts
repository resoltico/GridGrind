import info.solidsoft.gradle.pitest.PitestPluginExtension
import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoReport

plugins {
    `java-library`
    id("gridgrind.java-conventions")
    id("gridgrind.mutation-conventions")
}

description = "Canonical workbook model, vocabulary and Java authoring"

// Project dependencies keep runtime resolution under the owning project lock.
val jsonMutationTests = configurations.create("jsonMutationTests") {
    isCanBeConsumed = false
    isCanBeResolved = true
}

dependencies {
    jsonMutationTests(project(path = ":json", configuration = "mutationTests"))
}

tasks.named<info.solidsoft.gradle.pitest.PitestTask>("pitest") {
    additionalClasspath.from(jsonMutationTests)
}

configure<PitestPluginExtension> {
    targetClasses.set(
        setOf(
            "dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation*",
            "dev.erst.gridgrind.core.vocabulary.RgbHex",
            "dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdInput*",
            "dev.erst.gridgrind.core.model.validation.DataValidationRuleInput*",
            "dev.erst.gridgrind.core.catalog.FieldConstraint*",
            "dev.erst.gridgrind.core.catalog.CatalogStepTemplateDefaults*",
            "dev.erst.gridgrind.core.catalog.CatalogStepTemplateSupport*",
            "dev.erst.gridgrind.core.validation.WorkbookStaticMaterializationValidation*",
        ),
    )
    targetTests.set(
        setOf(
            "dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidationTest",
            "dev.erst.gridgrind.core.model.security.AdvancedMutationProtocolTypesTest",
            "dev.erst.gridgrind.core.model.validation.ConditionalFormattingDataBarBoundsTest",
            "dev.erst.gridgrind.core.model.security.ProtocolDefaultingCoverageTest",
            "dev.erst.gridgrind.core.model.validation.ProtocolResidualCoverageTest",
            "dev.erst.gridgrind.core.catalog.FieldConstraintTest",
            "dev.erst.gridgrind.core.catalog.CatalogStepTemplateSupportTest",
            "dev.erst.gridgrind.core.vocabulary.RgbHexTest",
            "dev.erst.gridgrind.json.codec.WireStringPropertyTest",
            "dev.erst.gridgrind.core.validation.WorkbookStaticMaterializationValidationTest",
            "dev.erst.gridgrind.core.validation.WorkbookStaticRequestContractTest",
        ),
    )
}

val downstreamCoverageProjects =
    listOf(project(":json"), project(":engine"), project(":cli"))

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
    // Jackson 3.x still owns annotations via the Jackson 2.x coordinates and package namespace.
    api(libs.jackson.annotations)
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
