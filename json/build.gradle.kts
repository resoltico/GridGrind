plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "JSON codecs and tolerant request intake"

dependencies {
    api(project(":core"))
    implementation(libs.jackson.annotations)
    implementation(libs.jackson.databind)
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.json.schema.validator)
    testImplementation(libs.poi.ooxml)
    testImplementation(libs.poi.ooxml.full)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}

tasks.withType<Test>().configureEach {
    systemProperty("gridgrind.root.dir", rootProject.projectDir.absolutePath)
    inputs.dir(rootProject.layout.projectDirectory.dir("examples")).withPropertyName("exampleRequests")
    inputs.dir(rootProject.layout.projectDirectory.dir("docs/reference/schema")).withPropertyName("schemaReferences")
}

// PIT discovers tests from class directories; include their production runtime as well.
configurations.create("mutationTests") {
    isCanBeConsumed = true
    isCanBeResolved = false
    extendsFrom(configurations.testImplementation.get(), configurations.testRuntimeOnly.get())
    attributes {
        attribute(Usage.USAGE_ATTRIBUTE, objects.named(Usage.JAVA_RUNTIME))
        attribute(Category.CATEGORY_ATTRIBUTE, objects.named(Category.LIBRARY))
        attribute(LibraryElements.LIBRARY_ELEMENTS_ATTRIBUTE, objects.named(LibraryElements.CLASSES))
    }
    listOf(sourceSets.main.get(), sourceSets.test.get()).forEach { sourceSet ->
        sourceSet.output.classesDirs.forEach { directory ->
            outgoing.artifact(directory) {
                builtBy(tasks.named(sourceSet.classesTaskName))
            }
        }
    }
}

// Generated references are updated explicitly; tests verify the committed artifacts independently.
tasks.register<JavaExec>("generateSchemas") {
    group = "documentation"
    description = "Writes request, grant and response JSON Schemas from the core model."
    classpath = sourceSets.main.get().runtimeClasspath
    mainClass.set("dev.erst.gridgrind.json.schema.JsonSchemas")
    args(rootProject.layout.projectDirectory.dir("docs/reference/schema").asFile.absolutePath)
    inputs.files(sourceSets.main.get().output)
    outputs.dir(rootProject.layout.projectDirectory.dir("docs/reference/schema"))
}

// Explicit regeneration and verification can share one invocation without racing on references.
tasks.named<Test>("test") {
    mustRunAfter("generateSchemas")
}
