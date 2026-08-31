import dev.erst.gridgrind.buildlogic.VerifyRuntimeLegalInventoryTask
import org.gradle.api.artifacts.component.ModuleComponentIdentifier
import org.gradle.api.file.DuplicatesStrategy

plugins {
    application
    id("gridgrind.java-conventions")
    id("gridgrind.mutation-conventions")
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
        "bcpkix-jdk18on-1.85.jar",
        "bcprov-jdk18on-1.85.2.jar",
        "bcutil-jdk18on-1.85.jar",
        "commons-codec-1.20.0.jar",
        "commons-collections4-4.5.0.jar",
        "commons-compress-1.28.0.jar",
        "commons-io-2.21.0.jar",
        "commons-lang3-3.18.0.jar",
        "commons-math3-3.6.1.jar",
        "curvesapi-1.08.jar",
        "jackson-annotations-2.22.jar",
        "jackson-core-3.2.2.jar",
        "jackson-databind-3.2.2.jar",
        "jakarta.activation-api-2.1.4.jar",
        "jakarta.xml.bind-api-4.0.5.jar",
        "log4j-api-2.26.1.jar",
        "log4j-core-2.26.1.jar",
        "log4j-slf4j2-impl-2.26.1.jar",
        "poi-5.5.1.jar",
        "poi-ooxml-5.5.1.jar",
        "poi-ooxml-full-5.5.1.jar",
        "poi-ooxml-lite-5.5.1.jar",
        "slf4j-api-2.0.17.jar",
        "stax2-api-4.2.2.jar",
        "woodstox-core-7.1.0.jar",
        "xmlbeans-5.3.0.jar",
        "xmlsec-4.0.4.jar",
    )
val reviewedNoticeMarkers =
    setOf(
        "Apache POI 5.5.1",
        "Apache XMLBeans 5.3.0",
        "Apache Log4j API 2.26.1",
        "Apache Santuario - XML Security for Java 4.0.4",
        "Apache Commons Codec 1.20.0",
        "Apache Commons Collections 4.5.0",
        "Apache Commons Compress 1.28.0",
        "Apache Commons IO 2.21.0",
        "Apache Commons Lang 3.18.0",
        "Apache Commons Math 3.6.1",
        "Apache POI CustomXMLMappings.xlsx test data",
        "Jackson Databind 3.2.2, Jackson Core 3.2.2, and Jackson Annotations 2.22",
        "SparseBitSet 1.3",
        "Woodstox Core 7.1.0",
        "Bouncy Castle (bcpkix-jdk18on 1.85, bcprov-jdk18on 1.85.2, bcutil-jdk18on 1.85)",
        "SLF4J API 2.0.17",
        "Stax2 API 4.2.2",
        "Jakarta Activation API 2.1.4",
        "Jakarta XML Binding API 4.0.5",
        "CurvesAPI 1.08",
        "Boost Software License, Version 1.0",
    )

description = "CLI transport adapter for the GridGrind protocol"

gridgrindMutation {
    targetClasses.set(
        setOf(
            "dev.erst.gridgrind.cli.CliLookupImmediateCommandParser*",
            "dev.erst.gridgrind.cli.GridGrindCliRecipeDiscoveryCommands*",
            "dev.erst.gridgrind.cli.examples.RecipeWorkspacePublisher*",
        ),
    )
    targetTests.set(
        setOf(
            "dev.erst.gridgrind.cli.CliArgumentsTest",
            "dev.erst.gridgrind.cli.GridGrindCliAssetRecipeMaterializationTest",
            "dev.erst.gridgrind.cli.GridGrindCliCatalogCommandTest",
            "dev.erst.gridgrind.cli.GridGrindCliRecipeDiscoveryContractTest",
            "dev.erst.gridgrind.cli.GridGrindCliRecipeTransportTest",
            "dev.erst.gridgrind.cli.GridGrindCommandErrorClassificationTest",
            "dev.erst.gridgrind.cli.examples.RecipeWorkspacePublisherTest",
        ),
    )
    mutationThreshold.set(100)
    coverageThreshold.set(100)
    testStrengthThreshold.set(100)
    maxSurviving.set(0)
}

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
    applicationDefaultJvmArgs = listOf("--enable-native-access=ALL-UNNAMED")
    mainClass = "dev.erst.gridgrind.cli.App"
}

distributions.named("shadow") {
    distributionBaseName.set("gridgrind")
    contents {
        from(packagedLegalFiles)
    }
}

val cleanShadowStartScripts = tasks.register<Delete>("cleanShadowStartScripts") {
    delete(layout.buildDirectory.dir("scriptsShadow"))
}

val cleanLegacyCliDistribution = tasks.register<Delete>("cleanLegacyCliDistribution") {
    delete(layout.buildDirectory.dir("scripts"))
    delete(layout.buildDirectory.dir("install/cli"))
    delete(layout.buildDirectory.dir("install/cli-shadow"))
}

val cleanInstalledShadowDist = tasks.register<Delete>("cleanInstalledShadowDist") {
    delete(layout.buildDirectory.dir("install/gridgrind"))
}

val cleanDistributionArchives = tasks.register<Delete>("cleanDistributionArchives") {
    delete(layout.buildDirectory.dir("distributions"))
}

tasks.named<JavaExec>("run") {
    workingDir = rootProject.projectDir
}

tasks.named<org.gradle.jvm.application.tasks.CreateStartScripts>("startScripts") {
    enabled = false
}

tasks.named<org.gradle.jvm.application.tasks.CreateStartScripts>("startShadowScripts") {
    applicationName = "gridgrind"
    dependsOn(cleanShadowStartScripts)
}

tasks.named<Sync>("installDist") {
    enabled = false
}

tasks.named<Sync>("installShadowDist") {
    dependsOn(cleanLegacyCliDistribution)
    dependsOn(cleanInstalledShadowDist)
}

tasks.named("distZip") {
    enabled = false
}

tasks.named("distTar") {
    enabled = false
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
    dependsOn(verifyRuntimeLegalInventory)
    archiveBaseName = "gridgrind"
    archiveVersion = ""
    archiveClassifier = ""
    isReproducibleFileOrder = true

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

    // Bundle the curated attribution notice and license texts into META-INF/.
    // LICENSE is the MIT license for GridGrind's own code.
    // The Apache, BSD 2-Clause, and component-specific BSD 3-Clause files cover all bundled
    // third-party license families not reproduced directly in NOTICE.
    packagedLegalFileNames.forEach { legalFileName ->
        from(rootProject.file(legalFileName)) { into("META-INF") }
    }

    manifest {
        attributes(
            "Enable-Native-Access" to "ALL-UNNAMED",
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
