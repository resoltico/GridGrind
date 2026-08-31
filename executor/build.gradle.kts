import org.gradle.api.plugins.quality.Pmd
import org.gradle.api.tasks.SourceSetContainer
import org.gradle.api.tasks.testing.Test
import org.gradle.kotlin.dsl.getByType
import org.gradle.kotlin.dsl.named
import org.gradle.kotlin.dsl.register
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport

plugins {
    `java-library`
    id("gridgrind.java-conventions")
    id("gridgrind.mutation-conventions")
}

description = "GridGrind parity and execution-regression verification surface"

val sourceSets = extensions.getByType<SourceSetContainer>()
val mainSourceSet = sourceSets.getByName("main")
val architectureTestSourceSet =
    sourceSets.create("architectureTest") {
        java.setSrcDirs(listOf("src/architectureTest/java"))
        resources.setSrcDirs(listOf("src/architectureTest/resources"))
        compileClasspath += mainSourceSet.output
        runtimeClasspath += mainSourceSet.runtimeClasspath
    }
val parityTestSourceSet =
    sourceSets.create("parityTest") {
        java.setSrcDirs(listOf("src/parityTest/java"))
        resources.setSrcDirs(listOf("src/parityTest/resources"))
        compileClasspath += mainSourceSet.output + configurations.getByName("testCompileClasspath")
        runtimeClasspath += output + compileClasspath
    }

configurations.named(parityTestSourceSet.implementationConfigurationName) {
    extendsFrom(configurations.getByName("testImplementation"))
}
configurations.named(parityTestSourceSet.runtimeOnlyConfigurationName) {
    extendsFrom(configurations.getByName("testRuntimeOnly"))
}

dependencies {
    implementation(libs.archunit.junit6)
    implementation(project(":authoring-java"))
    implementation(project(":cli"))
    implementation(project(":contract"))
    implementation(project(":engine"))
    implementation(project(":excel-foundation"))
    testImplementation(project(":engine"))
    testImplementation(project(":cli"))
    testImplementation(libs.archunit.junit6)
    testImplementation(libs.jackson.databind)
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testImplementation(libs.poi.ooxml.full)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
    add(architectureTestSourceSet.implementationConfigurationName, platform(libs.junit.bom))
    add(architectureTestSourceSet.implementationConfigurationName, libs.archunit.junit6)
    add(architectureTestSourceSet.implementationConfigurationName, libs.junit.jupiter)
    add(architectureTestSourceSet.runtimeOnlyConfigurationName, libs.junit.platform.launcher)
    add(parityTestSourceSet.implementationConfigurationName, project(":engine"))
    add(parityTestSourceSet.implementationConfigurationName, testFixtures(project(":engine")))
    add(parityTestSourceSet.implementationConfigurationName, libs.jackson.databind)
    add(parityTestSourceSet.implementationConfigurationName, libs.junit.jupiter)
    add(parityTestSourceSet.implementationConfigurationName, libs.poi.ooxml)
    add(parityTestSourceSet.implementationConfigurationName, libs.poi.ooxml.full)
    add(parityTestSourceSet.implementationConfigurationName, libs.bcpkix)
    add(parityTestSourceSet.implementationConfigurationName, libs.bcprov)
    add(parityTestSourceSet.implementationConfigurationName, libs.bcutil)
    add(parityTestSourceSet.implementationConfigurationName, libs.xmlsec)
    add(parityTestSourceSet.runtimeOnlyConfigurationName, libs.slf4j.api)
    add(parityTestSourceSet.runtimeOnlyConfigurationName, libs.junit.platform.launcher)
    add(parityTestSourceSet.runtimeOnlyConfigurationName, libs.log4j.core)
}

val architectureTest =
    tasks.register<Test>("architectureTest") {
        description = "Runs the mandatory ArchUnit-engine product architecture rules."
        group = "verification"
        testClassesDirs = architectureTestSourceSet.output.classesDirs
        classpath = architectureTestSourceSet.runtimeClasspath
        useJUnitPlatform {
            includeEngines("archunit")
        }
        failOnNoDiscoveredTests = true
        shouldRunAfter(tasks.named<Test>("test"))
    }

val parityTest =
    tasks.register<Test>("parityTest") {
        description = "Runs the dedicated Apache POI XSSF parity verification suite."
        group = "verification"
        testClassesDirs = parityTestSourceSet.output.classesDirs
        classpath = parityTestSourceSet.runtimeClasspath
        useJUnitPlatform()
        shouldRunAfter(tasks.named<Test>("test"))
    }

val architectureTestContract =
    tasks.register<Test>("architectureTestContract") {
        description = "Verifies the architecture gate's effective runtime configuration and rule inventory."
        group = "verification"
        testClassesDirs = architectureTestSourceSet.output.classesDirs
        classpath = architectureTestSourceSet.runtimeClasspath
        useJUnitPlatform {
            includeEngines("junit-jupiter")
        }
        failOnNoDiscoveredTests = true
        shouldRunAfter(architectureTest)
    }

fun parityCoverageExecutionData() =
    provider {
        listOf(
            parityTest.get().extensions.getByType(JacocoTaskExtension::class.java).destinationFile
        )
    }

tasks.named("check") {
    dependsOn(architectureTest)
    dependsOn(architectureTestContract)
    dependsOn(parityTest)
    dependsOn("pmdArchitectureTest")
    dependsOn("pmdParityTest")
}

tasks.named<JacocoReport>("jacocoTestReport") {
    dependsOn(parityTest)
    executionData.from(parityCoverageExecutionData())
}

tasks.named<JacocoCoverageVerification>("jacocoTestCoverageVerification") {
    dependsOn(parityTest)
    executionData.from(parityCoverageExecutionData())
}

tasks.named<Pmd>("pmdParityTest") {
    ruleSetFiles = files(rootProject.file("gradle/pmd/test-ruleset.xml"))
    ruleSets = emptyList()
}

tasks.named<Pmd>("pmdArchitectureTest") {
    ruleSetFiles = files(rootProject.file("gradle/pmd/test-ruleset.xml"))
    ruleSets = emptyList()
}

gridgrindMutation {
    targetClasses.set(
        setOf(
            "dev.erst.gridgrind.architecture.GridGrindProductLocations*",
            "dev.erst.gridgrind.architecture.EngineImplementationTypeClassifier*",
            "dev.erst.gridgrind.architecture.ExportedApiImplementationTypes*",
            "dev.erst.gridgrind.architecture.ProductArchitectureRules*",
            "dev.erst.gridgrind.architecture.ProductDependencyArchitectureRules*",
            "dev.erst.gridgrind.architecture.ProductDomainShapeArchitectureRules*",
            "dev.erst.gridgrind.architecture.ProductToolingSeamArchitectureRules*",
        ),
    )
    targetTests.set(
        setOf(
            "dev.erst.gridgrind.architecture.GridGrindProductLocationsTest",
            "dev.erst.gridgrind.architecture.ProductArchitectureRuleConstructionTest",
            "dev.erst.gridgrind.architecture.ProductArchitectureRuleInventoryTest",
            "dev.erst.gridgrind.architecture.ProductDependencyArchitectureRulesTest",
            "dev.erst.gridgrind.architecture.ProductDomainShapeArchitectureRulesTest",
            "dev.erst.gridgrind.architecture.ProductToolingSeamArchitectureRulesTest",
            "dev.erst.gridgrind.architecture.ProductArchitectureRulesIntegrationTest",
        ),
    )
    mutationThreshold.set(100)
    coverageThreshold.set(100)
    testStrengthThreshold.set(100)
    maxSurviving.set(0)
}

tasks.named("pitest") {
    mustRunAfter(":engine:pitest")
}
