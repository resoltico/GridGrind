plugins {
    application
    id("gridgrind.java-conventions")
    alias(libs.plugins.shadow)
}

description = "CLI transport adapter for the GridGrind protocol"

dependencies {
    implementation(project(":executor"))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testRuntimeOnly(libs.junit.platform.launcher)
    runtimeOnly(libs.log4j.core)
    runtimeOnly(libs.log4j.slf4j2.impl)
}

application {
    mainClass = "dev.erst.gridgrind.cli.App"
}

tasks.named<JavaExec>("run") {
    workingDir = rootProject.projectDir
}

tasks.named<Test>("test") {
    jvmArgs("--enable-native-access=ALL-UNNAMED")
}

tasks.named<com.github.jengelman.gradle.plugins.shadow.tasks.ShadowJar>("shadowJar") {
    archiveBaseName = "gridgrind"
    archiveVersion = ""
    archiveClassifier = ""

    // Merge ServiceLoader registrations from all bundled JARs.
    mergeServiceFiles()

    // Exclude per-dependency META-INF license and notice files to prevent conflicts
    // and silent overwrites. GridGrind bundles its own curated NOTICE, MIT LICENSE,
    // and the Apache License 2.0 text that covers all bundled Apache-licensed components.
    exclude("META-INF/LICENSE", "META-INF/LICENSE.txt", "META-INF/LICENSE.md")
    exclude("META-INF/NOTICE", "META-INF/NOTICE.txt", "META-INF/NOTICE.md")
    exclude("META-INF/DEPENDENCIES")

    // Bundle the curated attribution notice and both license texts into META-INF/.
    // NOTICE covers Apache POI, Log4j Core, and Jackson Databind attribution.
    // LICENSE is the MIT license for GridGrind's own code.
    // LICENSE-APACHE-2.0 satisfies Apache License 2.0 Section 4(a) for bundled components.
    from(rootProject.file("NOTICE")) { into("META-INF") }
    from(rootProject.file("PATENTS.md")) { into("META-INF") }
    from(rootProject.file("LICENSE")) { into("META-INF") }
    from(rootProject.file("LICENSE-APACHE-2.0")) { into("META-INF") }
    from(rootProject.file("LICENSE-BSD-3-CLAUSE")) { into("META-INF") }

    manifest {
        attributes(
            "Enable-Native-Access" to "ALL-UNNAMED",
            "Implementation-Title" to "GridGrind",
            "Implementation-Version" to project.version,
            "Implementation-Vendor" to "Ervins Strauhmanis",
            "Implementation-License" to "MIT",
        )
    }
}

tasks.named<ProcessResources>("processResources") {
    val description: String = providers.gradleProperty("gridgrindDescription").get()
    filesMatching("gridgrind.properties") {
        expand(mapOf("gridgrindDescription" to description))
    }
}

tasks.named("assemble") {
    dependsOn("shadowJar")
}
