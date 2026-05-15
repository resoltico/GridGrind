plugins {
    application
    id("gridgrind.java-conventions")
    alias(libs.plugins.shadow)
}

description = "CLI transport adapter for the GridGrind protocol"

dependencies {
    implementation(project(":contract"))
    implementation(project(":engine"))
    implementation(project(":excel-foundation"))
    implementation(libs.jackson.databind)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testRuntimeOnly(libs.junit.platform.launcher)
    runtimeOnly(libs.log4j.core)
    runtimeOnly(libs.log4j.slf4j2.impl)
}

application {
    applicationDefaultJvmArgs = listOf("--enable-native-access=ALL-UNNAMED")
    mainClass = "dev.erst.gridgrind.cli.App"
}

val cleanStartScripts by tasks.registering(Delete::class) {
    delete(layout.buildDirectory.dir("scripts"))
}

val cleanShadowStartScripts by tasks.registering(Delete::class) {
    delete(layout.buildDirectory.dir("scriptsShadow"))
}

val cleanInstalledDist by tasks.registering(Delete::class) {
    delete(layout.buildDirectory.dir("install/cli"))
}

val cleanInstalledShadowDist by tasks.registering(Delete::class) {
    delete(layout.buildDirectory.dir("install/cli-shadow"))
}

val cleanDistributionArchives by tasks.registering(Delete::class) {
    delete(layout.buildDirectory.dir("distributions"))
}

tasks.named<JavaExec>("run") {
    workingDir = rootProject.projectDir
}

tasks.named<org.gradle.jvm.application.tasks.CreateStartScripts>("startScripts") {
    applicationName = "gridgrind"
    dependsOn(cleanStartScripts)
}

tasks.named<org.gradle.jvm.application.tasks.CreateStartScripts>("startShadowScripts") {
    applicationName = "gridgrind"
    dependsOn(cleanShadowStartScripts)
}

tasks.named<Sync>("installDist") {
    dependsOn(cleanInstalledDist)
}

tasks.named<Sync>("installShadowDist") {
    dependsOn(cleanInstalledShadowDist)
}

tasks.named("distZip") {
    dependsOn(cleanDistributionArchives)
}

tasks.named("distTar") {
    dependsOn(cleanDistributionArchives)
}

tasks.named("shadowDistZip") {
    dependsOn(cleanDistributionArchives)
}

tasks.named("shadowDistTar") {
    dependsOn(cleanDistributionArchives)
}

tasks.named<Test>("test") {
    jvmArgs("--enable-native-access=ALL-UNNAMED")
}

tasks.named<com.github.jengelman.gradle.plugins.shadow.tasks.ShadowJar>("shadowJar") {
    archiveBaseName = "gridgrind"
    archiveVersion = ""
    archiveClassifier = ""
    isReproducibleFileOrder = true

    // Merge ServiceLoader registrations from all bundled JARs.
    mergeServiceFiles()

    // Exclude per-dependency META-INF license and notice files to prevent conflicts
    // and silent overwrites. GridGrind bundles its own curated NOTICE plus the
    // project-owned license texts for all bundled dependency license families.
    exclude("META-INF/LICENSE", "META-INF/LICENSE.txt", "META-INF/LICENSE.md")
    exclude("META-INF/NOTICE", "META-INF/NOTICE.txt", "META-INF/NOTICE.md")
    exclude("META-INF/DEPENDENCIES")

    // Bundle the curated attribution notice and license texts into META-INF/.
    // LICENSE is the MIT license for GridGrind's own code.
    // LICENSE-APACHE-2.0, LICENSE-BSD-2-CLAUSE, LICENSE-BSD-3-CLAUSE, and LICENSE-EDL-1.0 cover
    // bundled third-party components under those respective license families.
    from(rootProject.file("NOTICE")) { into("META-INF") }
    from(rootProject.file("PATENTS.md")) { into("META-INF") }
    from(rootProject.file("LICENSE")) { into("META-INF") }
    from(rootProject.file("LICENSE-APACHE-2.0")) { into("META-INF") }
    from(rootProject.file("LICENSE-BSD-2-CLAUSE")) { into("META-INF") }
    from(rootProject.file("LICENSE-BSD-3-CLAUSE")) { into("META-INF") }
    from(rootProject.file("LICENSE-EDL-1.0")) { into("META-INF") }

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
    val version: String = providers.gradleProperty("version").get()
    inputs.property("gridgrindDescription", description)
    inputs.property("gridgrindVersion", version)
    filesMatching("gridgrind.properties") {
        expand(
            mapOf(
                "gridgrindDescription" to description,
                "gridgrindVersion" to version,
            )
        )
    }
}

pluginManager.withPlugin("java") {
    tasks.register<JavaExec>("writeRepositoryExamples") {
        group = "documentation"
        description =
            "Regenerates the checkout-rooted examples/*.json fixtures from the CLI-owned example registry."
        dependsOn(tasks.named("testClasses"))
        classpath =
            project.the<org.gradle.api.plugins.JavaPluginExtension>()
                .sourceSets
                .named("test")
                .get()
                .runtimeClasspath
        mainClass = "dev.erst.gridgrind.cli.discovery.ExampleRequestFixturesWriter"
        workingDir = rootProject.projectDir
    }
}

tasks.named("assemble") {
    dependsOn("shadowJar")
}
