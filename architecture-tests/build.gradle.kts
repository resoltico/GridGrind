import org.gradle.api.plugins.quality.Pmd
import org.gradle.api.tasks.ClasspathNormalizer
import org.gradle.api.tasks.SourceSetContainer
import org.gradle.api.tasks.testing.Test
import java.io.File

plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "Product architecture rules and their focused negative controls"

val productProjects = rootProject.subprojects.filter { it != project }.sortedBy { it.path }
val sourceSets = extensions.getByType<SourceSetContainer>()
val mainSourceSet = sourceSets.getByName("main")
val architectureSourceSet = sourceSets.create("architectureTest") {
    compileClasspath += mainSourceSet.output
    runtimeClasspath += mainSourceSet.runtimeClasspath
}

dependencies {
    implementation(libs.archunit.junit6)
    productProjects.forEach { implementation(project(it.path)) }
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.jackson.databind)
    testImplementation(libs.poi.ooxml)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
    add(architectureSourceSet.implementationConfigurationName, platform(libs.junit.bom))
    add(architectureSourceSet.implementationConfigurationName, libs.archunit.junit6)
    add(architectureSourceSet.runtimeOnlyConfigurationName, libs.junit.platform.launcher)
    add(architectureSourceSet.runtimeOnlyConfigurationName, libs.log4j.core)
}

val architectureTest = tasks.register<Test>("architectureTest") {
    description = "Runs architecture rules over the compiled product classes."
    group = "verification"
    testClassesDirs = architectureSourceSet.output.classesDirs
    classpath = architectureSourceSet.runtimeClasspath
    useJUnitPlatform { includeEngines("archunit") }
    failOnNoDiscoveredTests = true
    shouldRunAfter(tasks.named<Test>("test"))
}

gradle.projectsEvaluated {
    val classRoots = productProjects.flatMap {
        it.extensions.getByType<SourceSetContainer>().getByName("main").output.classesDirs.files
    }.sortedBy { it.absolutePath }
    architectureTest.configure {
        dependsOn(productProjects.map { "${it.path}:classes" })
        inputs.files(classRoots).withPropertyName("productClasses").withNormalizer(ClasspathNormalizer::class.java)
        systemProperty("gridgrind.architecture.classRoots", classRoots.joinToString(File.pathSeparator) { it.absolutePath })
    }
}

tasks.named<Pmd>("pmdArchitectureTest") {
    ruleSetFiles = files(rootProject.file("gradle/pmd/test-ruleset.xml"))
    ruleSets = emptyList()
}

tasks.named("check") {
    dependsOn(architectureTest)
    dependsOn("pmdArchitectureTest")
}
