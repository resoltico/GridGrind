import dev.erst.gridgrind.buildlogic.coverageTestTasks
import info.solidsoft.gradle.pitest.PitestPluginExtension
import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoReport

plugins {
    `java-library`
    `java-test-fixtures`
    id("gridgrind.java-conventions")
    id("gridgrind.mutation-conventions")
}

description = "Core GridGrind workbook automation and request execution engine"

configure<PitestPluginExtension> {
    targetClasses.set(
        setOf(
            "dev.erst.gridgrind.engine.runtime.FormulaOriginTracker*",
            "dev.erst.gridgrind.engine.runtime.RequestPreflightPaths*",
            "dev.erst.gridgrind.engine.validation.DataValidationAdapter*",
        ),
    )
    targetTests.set(
        setOf(
            "dev.erst.gridgrind.engine.runtime.FormulaOriginTrackerTest",
            "dev.erst.gridgrind.engine.runtime.RequestPreflightTest",
            "dev.erst.gridgrind.engine.runtime.RequestPreflightFormulaEnvironmentTest",
            "dev.erst.gridgrind.engine.validation.ValidationComparisonOperatorsTest",
        ),
    )
}

tasks.named("pitest") {
    mustRunAfter(":core:pitest")
}

// Engine behavior is exercised through both runtime adapters; coverage evidence must include both.
val downstreamCoverageProjects = listOf(project(":json"), project(":cli"))

fun downstreamCoverageTaskPaths(): List<String> =
    downstreamCoverageProjects.flatMap { downstreamProject ->
        downstreamProject.coverageTestTasks().map { task -> task.path }
    }

fun downstreamCoverageExecutionData() =
    provider {
        downstreamCoverageProjects.flatMap { downstreamProject ->
            downstreamProject.coverageTestTasks().map { testTask ->
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
    api(project(":core"))
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

val paritySourceSet = the<org.gradle.api.tasks.SourceSetContainer>().create("parityTest") {
    compileClasspath += sourceSets.main.get().output + configurations.getByName("testCompileClasspath")
    runtimeClasspath += output + compileClasspath
}
configurations.named(paritySourceSet.implementationConfigurationName) {
    extendsFrom(configurations.getByName("testImplementation"))
}
configurations.named(paritySourceSet.runtimeOnlyConfigurationName) {
    extendsFrom(configurations.getByName("testRuntimeOnly"))
}
dependencies {
    testImplementation(project(":json"))
    testImplementation(project(":cli"))
    testImplementation(libs.jackson.databind)
    add(paritySourceSet.implementationConfigurationName, libs.jackson.databind)
    add(paritySourceSet.implementationConfigurationName, libs.bcpkix)
    add(paritySourceSet.implementationConfigurationName, libs.bcprov)
    add(paritySourceSet.implementationConfigurationName, libs.bcutil)
    add(paritySourceSet.implementationConfigurationName, libs.xmlsec)
    add(paritySourceSet.runtimeOnlyConfigurationName, libs.slf4j.api)
}
val parityTest = tasks.register<Test>("parityTest") {
    inputs.files(
        rootProject.fileTree("docs") { include("*.md") },
        rootProject.file("README.md"),
    ).withPathSensitivity(PathSensitivity.RELATIVE)
    description = "Runs the XSSF behavioural parity corpus."
    group = "verification"
    testClassesDirs = paritySourceSet.output.classesDirs
    classpath = paritySourceSet.runtimeClasspath
    useJUnitPlatform()
    failOnNoDiscoveredTests = true
    shouldRunAfter(tasks.named<Test>("test"))
}
tasks.named<org.gradle.api.plugins.quality.Pmd>("pmdParityTest") {
    ruleSetFiles = files(rootProject.file("gradle/pmd/test-ruleset.xml"))
    ruleSets = emptyList()
}
tasks.named("check") {
    dependsOn(parityTest, "pmdParityTest")
}
