import dev.erst.gridgrind.buildlogic.VerifyRuntimeLegalInventoryTask
import org.gradle.api.artifacts.component.ModuleComponentIdentifier
import org.gradle.api.file.DuplicatesStrategy

plugins {
    application
    id("gridgrind.java-conventions")
    alias(libs.plugins.shadow)
}

val distributions = the<org.gradle.api.distribution.DistributionContainer>()
val packagedLegalFileNames =
    listOf(
        "LICENSE",
        "NOTICE",
        "PATENTS.md",
        "LICENSE-APACHE-2.0",
        "LICENSE-BSD-2-CLAUSE",
        "LICENSE-BSD-3-CLAUSE",
    )
val packagedLegalFiles = files(packagedLegalFileNames.map(rootProject::file))
val reviewedRuntimeArtifactNames =
    setOf(
        "SparseBitSet-1.3.jar",
        "bcpkix-jdk18on-1.86.jar",
        "bcprov-jdk18on-1.86.jar",
        "bcutil-jdk18on-1.86.jar",
        "commons-codec-1.23.0.jar",
        "commons-collections4-4.6.0.jar",
        "commons-compress-1.28.0.jar",
        "commons-io-2.22.0.jar",
        "commons-lang3-3.21.0.jar",
        "commons-math3-3.6.1.jar",
        "curvesapi-1.08.jar",
        "jackson-annotations-2.22.jar",
        "jackson-core-3.2.4.jar",
        "jackson-databind-3.2.4.jar",
        "jakarta.activation-api-2.1.4.jar",
        "jakarta.xml.bind-api-4.0.5.jar",
        "log4j-api-2.26.1.jar",
        "log4j-core-2.26.1.jar",
        "log4j-slf4j2-impl-2.26.1.jar",
        "poi-6.0.0.jar",
        "poi-ooxml-6.0.0.jar",
        "poi-ooxml-full-6.0.0.jar",
        "poi-ooxml-lite-6.0.0.jar",
        "slf4j-api-2.0.20.jar",
        "stax2-api-4.3.2.jar",
        "woodstox-core-7.3.0.jar",
        "xmlbeans-5.4.1.jar",
        "xmlsec-4.0.4.jar",
    )
val reviewedNoticeMarkers =
    setOf(
        "Apache POI 6.0.0",
        "Apache XMLBeans 5.4.1",
        "Apache Log4j API 2.26.1",
        "Apache Santuario - XML Security for Java 4.0.4",
        "Apache Commons Codec 1.23.0",
        "Apache Commons Collections 4.6.0",
        "Apache Commons Compress 1.28.0",
        "Apache Commons IO 2.22.0",
        "Apache Commons Lang 3.21.0",
        "Apache Commons Math 3.6.1",
        "Apache POI CustomXMLMappings.xlsx test data",
        "Jackson Databind 3.2.4, Jackson Core 3.2.4, and Jackson Annotations 2.22",
        "SparseBitSet 1.3",
        "Woodstox Core 7.3.0",
        "Woodstox 7.3.0 Shaded Components",
        "ISO-RELAX 20090621",
        "RELAX NG datatype API 20020414",
        "The Apache Software License, Version 1.1",
        "Bouncy Castle (bcpkix-jdk18on 1.86, bcprov-jdk18on 1.86, bcutil-jdk18on 1.86)",
        "SLF4J API 2.0.20",
        "Stax2 API 4.3.2",
        "Jakarta Activation API 2.1.4",
        "Jakarta XML Binding API 4.0.5",
        "CurvesAPI 1.08",
        "Boost Software License, Version 1.0",
    )

description = "CLI transport adapter for the GridGrind protocol"


dependencies {
    testImplementation(testFixtures(project(":core")))
    implementation(project(":json"))
    implementation(project(":engine"))
    implementation(libs.jackson.databind)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testRuntimeOnly(libs.junit.platform.launcher)
    runtimeOnly(libs.log4j.core)
    runtimeOnly(libs.log4j.slf4j2.impl)
}

val runtimeClasspathConfiguration = configurations.named("runtimeClasspath")
val externalRuntimeArtifacts =
    runtimeClasspathConfiguration.get().incoming.artifactView {
        componentFilter { component -> component is ModuleComponentIdentifier }
    }.files
val verifyRuntimeLegalInventory =
    tasks.register<VerifyRuntimeLegalInventoryTask>("verifyRuntimeLegalInventory") {
        group = "verification"
        description = "Fails when the packaged runtime graph or its audited NOTICE inventory drifts"
        runtimeArtifacts.from(externalRuntimeArtifacts)
        auditedRuntimeArtifactNames.set(reviewedRuntimeArtifactNames)
        auditedNoticeMarkers.set(reviewedNoticeMarkers)
        noticeFile.set(rootProject.layout.projectDirectory.file("NOTICE"))
    }

tasks.named("check") {
    dependsOn(verifyRuntimeLegalInventory)
}

application {
    mainClass = "dev.erst.gridgrind.cli.App"
}

distributions.named("shadow") {
    distributionBaseName.set("gridgrind")
    contents {
        from(packagedLegalFiles)
    }
}

tasks.named<JavaExec>("run") {
    workingDir = rootProject.projectDir
}

tasks.named<org.gradle.jvm.application.tasks.CreateStartScripts>("startScripts") {
    enabled = false
}

tasks.named<org.gradle.jvm.application.tasks.CreateStartScripts>("startShadowScripts") {
    applicationName = "gridgrind"
}

tasks.named<Sync>("installDist") {
    enabled = false
}

tasks.named("distZip") {
    enabled = false
}

tasks.named("distTar") {
    enabled = false
}

tasks.named<Test>("test") {
    testLogging.exceptionFormat = org.gradle.api.tasks.testing.logging.TestExceptionFormat.FULL
    dependsOn("shadowJar", "installShadowDist", "shadowDistZip", "shadowDistTar")
    inputs.files(tasks.named("shadowJar"), tasks.named("installShadowDist"), tasks.named("shadowDistZip"), tasks.named("shadowDistTar"))
        .withPropertyName("deliveredArtifacts")
    inputs.dir(rootProject.layout.projectDirectory.dir("examples")).withPropertyName("repositoryExamples")
    inputs.files(rootProject.file("README.md"), rootProject.file("docs/EXECUTION_MODEL.md"), rootProject.file("docs/QUICK_START.md"))
        .withPropertyName("hostGuidance")
    useJUnitPlatform { excludeTags("container", "devcontainer", "documentation") }
    systemProperty("gridgrind.repository", rootProject.projectDir.absolutePath)
    systemProperty("gridgrind.golden.update", providers.gradleProperty("golden.update").getOrElse("false"))
}

tasks.named<com.github.jengelman.gradle.plugins.shadow.tasks.ShadowJar>("shadowJar") {
    dependsOn(verifyRuntimeLegalInventory)
    archiveBaseName = "gridgrind"
    archiveVersion = ""
    archiveClassifier = ""

    // Merge ServiceLoader registrations from all bundled JARs.
    filesMatching("META-INF/services/**") {
        duplicatesStrategy = DuplicatesStrategy.INCLUDE
    }
    mergeServiceFiles()
    failOnDuplicateEntries = true

    // Exclude per-dependency META-INF license and notice files to prevent conflicts
    // and silent overwrites. GridGrind bundles its own curated NOTICE plus the
    // project-owned license texts for all bundled dependency license families.
    exclude("META-INF/LICENSE", "META-INF/LICENSE.txt", "META-INF/LICENSE.md")
    exclude("META-INF/NOTICE", "META-INF/NOTICE.txt", "META-INF/NOTICE.md")
    exclude("META-INF/DEPENDENCIES")
    // These unused shaded build/interactive-tool artifacts have separate, ambiguous terms.
    // Keep runtime providers and Maven provenance properties; neither needs these entries.
    exclude("com/ctc/wstx/shaded/msv_core/datatype/xsd/CommandLineTester.class")
    exclude("META-INF/maven/com.sun.xml.bind.jaxb/isorelax/pom.xml")

    // Bundle the curated attribution notice and license texts into META-INF/.
    // LICENSE is the MPL 2.0 license for GridGrind's own code.
    // The Apache, BSD 2-Clause, and component-specific BSD 3-Clause files cover all bundled
    // third-party license families not reproduced directly in NOTICE.
    packagedLegalFileNames.forEach { legalFileName ->
        from(rootProject.file(legalFileName)) { into("META-INF") }
    }

    manifest {
        attributes(
            "Implementation-Title" to "GridGrind",
            "Implementation-Version" to project.version,
            "Implementation-Vendor" to "Ervins Strauhmanis",
            "Implementation-License" to "Multiple; see META-INF/NOTICE",
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
    from(rootProject.file("examples")) {
        include("*-assets/**")
        into("gridgrind/recipe-assets")
    }
    from(packagedLegalFiles) {
        into("licenses")
    }
}

pluginManager.withPlugin("java") {
    tasks.register<JavaExec>("writeRepositoryExamples") {
        group = "documentation"
        description =
            "Regenerates the checkout-rooted examples/*.json fixtures from the core-owned recipe registry."
        dependsOn(tasks.named("testClasses"))
        classpath =
            project.the<org.gradle.api.plugins.JavaPluginExtension>()
                .sourceSets
                .named("test")
                .get()
                .runtimeClasspath
        mainClass = "dev.erst.gridgrind.cli.io.ExampleRequestFixturesWriter"
        workingDir = rootProject.projectDir
    }
}

tasks.named("assemble") {
    dependsOn("shadowJar")
}

val hostTestSources = the<org.gradle.api.plugins.JavaPluginExtension>().sourceSets.named("test")
for (suite in listOf("container", "devcontainer")) {
    tasks.register<Test>(suite + "Test") {
        group = "verification"
        description = "Explicit Docker-backed " + suite + " verification."
        dependsOn("testClasses", "shadowJar")
        testClassesDirs = hostTestSources.get().output.classesDirs
        classpath = hostTestSources.get().runtimeClasspath
        useJUnitPlatform { includeTags(suite) }
        systemProperty("gridgrind.repository", rootProject.projectDir.absolutePath)
        extensions.configure<org.gradle.testing.jacoco.plugins.JacocoTaskExtension> { isEnabled = false }
        outputs.upToDateWhen { false }
        outputs.cacheIf { false }
    }
}

for (operation in listOf("generateReferenceDocs", "verifyGeneratedDocs")) {
    tasks.register<JavaExec>(operation) {
        group = "documentation"
        description = if (operation == "verifyGeneratedDocs") "Verifies reference documents and schemas against core metadata." else "Regenerates metadata-owned Markdown references."
        dependsOn("testClasses")
        classpath = hostTestSources.get().runtimeClasspath
        mainClass.set("dev.erst.gridgrind.cli.io.ReferenceDocumentsWriter")
        args(rootProject.projectDir.absolutePath, if (operation == "verifyGeneratedDocs") "--verify" else "--write")
    }
}

val documentationTest = tasks.register<Test>("documentationTest") {
    group = "verification"
    description = "Validates handwritten JSON examples and local documentation links."
    dependsOn("testClasses")
    testClassesDirs = hostTestSources.get().output.classesDirs
    classpath = hostTestSources.get().runtimeClasspath
    useJUnitPlatform { includeTags("documentation") }
    systemProperty("gridgrind.repository", rootProject.projectDir.absolutePath)
    inputs.dir(rootProject.layout.projectDirectory.dir("docs")).withPropertyName("documentation")
    inputs.files(rootProject.file("README.md"), rootProject.file("CHANGELOG.md"), rootProject.file("CHANGELOG_ARCHIVE.md"), rootProject.file("PATENTS.md"), rootProject.file("core/VOCABULARY.md"))
        .withPropertyName("entryDocuments")
    extensions.configure<org.gradle.testing.jacoco.plugins.JacocoTaskExtension> { isEnabled = false }
}
tasks.named("check") { dependsOn(documentationTest) }
