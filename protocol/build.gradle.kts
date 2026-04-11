import org.gradle.api.plugins.quality.Pmd
import org.gradle.api.tasks.SourceSetContainer
import org.gradle.api.tasks.testing.Test
import org.gradle.kotlin.dsl.getByType
import org.gradle.kotlin.dsl.named
import org.gradle.kotlin.dsl.register

plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "Structured GridGrind protocol and request execution layer"

val sourceSets = extensions.getByType<SourceSetContainer>()
val mainSourceSet = sourceSets.getByName("main")
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
    implementation(project(":engine"))
    api(libs.jackson.databind)
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testImplementation(libs.poi.ooxml.full)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
    add(parityTestSourceSet.implementationConfigurationName, testFixtures(project(":engine")))
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

val parityTest =
    tasks.register<Test>("parityTest") {
        description = "Runs the dedicated Apache POI XSSF parity verification suite."
        group = "verification"
        testClassesDirs = parityTestSourceSet.output.classesDirs
        classpath = parityTestSourceSet.runtimeClasspath
        useJUnitPlatform()
        shouldRunAfter(tasks.named<Test>("test"))
    }

tasks.named("check") {
    dependsOn(parityTest)
    dependsOn("pmdParityTest")
}

tasks.named<Pmd>("pmdParityTest") {
    ruleSetFiles = files(rootProject.file("gradle/pmd/test-ruleset.xml"))
    ruleSets = emptyList()
}
