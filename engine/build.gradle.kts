import dev.erst.gridgrind.buildlogic.NativeProviderModuleArguments
import dev.erst.gridgrind.buildlogic.NativeProviderVariants
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

// These compiler views are directories, applied to the unchanged named vendor modules.
val poiNativeCompile = configurations.create("poiNativeCompile") {
    attributes.attribute(NativeProviderVariants.REPAIR_MODULE, NativeProviderVariants.POI_MODULE)
    isCanBeConsumed = false
    isCanBeResolved = true
}
val ooxmlNativeCompile = configurations.create("ooxmlNativeCompile") {
    attributes.attribute(NativeProviderVariants.REPAIR_MODULE, NativeProviderVariants.OOXML_MODULE)
    isCanBeConsumed = false
    isCanBeResolved = true
}
dependencies {
    add(poiNativeCompile.name, project(path = ":native-provider", configuration = "poiCompileElements"))
    add(ooxmlNativeCompile.name, project(path = ":native-provider", configuration = "ooxmlCompileElements"))
    compileOnly(project(path = ":native-provider", configuration = "nativeCompileElements"))
    runtimeOnly(project(path = ":native-provider", configuration = "nativeRuntimeElements"))
}
val nativePatchArguments = objects.newInstance(NativeProviderModuleArguments::class.java).apply {
    poiClasses.from(poiNativeCompile)
    ooxmlClasses.from(ooxmlNativeCompile)
}
tasks.named<JavaCompile>("compileJava") {
    dependsOn(poiNativeCompile, ooxmlNativeCompile)
    options.compilerArgumentProviders.add(nativePatchArguments)
}
val nativeOwnerClasses = files(poiNativeCompile, ooxmlNativeCompile)
// Native owner tests use the declared classpath view. Public realm tests use delivery resources only.
sourceSets.matching { it.name != "main" }.all {
    val ownerSourceSet = this
    tasks.named<JavaCompile>(compileJavaTaskName) {
        classpath = nativeOwnerClasses + ownerSourceSet.compileClasspath
    }
}
tasks.withType<Test>().configureEach {
    if (name != "publicNativeTest" && name != "publicNativeLifetimeTest") {
        classpath = nativeOwnerClasses + classpath
        useJUnitPlatform { excludeTags("native-public", "native-lifetime") }
    }
}
val publicNativeTest = tasks.register<Test>("publicNativeTest") {
    group = "verification"
    description = "Verifies public native realms using stock parent JARs and the resource capsule"
    failOnNoDiscoveredTests = true
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    useJUnitPlatform {
        includeTags("native-public")
        excludeTags("native-lifetime")
    }
    shouldRunAfter(tasks.named<Test>("test"))
}

val publicNativeLifetimeTest = tasks.register<Test>("publicNativeLifetimeTest") {
    group = "verification"
    description = "Verifies public realm reclamation after JVM soft adaptation caches are cleared"
    failOnNoDiscoveredTests = true
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    useJUnitPlatform { includeTags("native-lifetime") }
    filter {
        includeTestsMatching("*NativePublicExecutionTest.realmsAreReclaimedWhenJvmSoftAdaptationCachesAreCleared")
        isFailOnNoMatchingTests = true
    }
    jvmArgs("-XX:SoftRefLRUPolicyMSPerMB=0")
    shouldRunAfter(publicNativeTest)
}

tasks.named("check") { dependsOn(publicNativeTest, publicNativeLifetimeTest) }

configure<PitestPluginExtension> {
    targetClasses.convention(
        setOf(
            "dev.erst.gridgrind.engine.runtime.FormulaOriginTracker*",
            "dev.erst.gridgrind.engine.runtime.RequestPreflightPaths*",
            "dev.erst.gridgrind.engine.validation.DataValidationAdapter*",
        ),
    )
    targetTests.convention(
        setOf(
            "dev.erst.gridgrind.engine.runtime.FormulaOriginTrackerTest",
            "dev.erst.gridgrind.engine.runtime.RequestPreflightTest",
            "dev.erst.gridgrind.engine.runtime.RequestPreflightFormulaEnvironmentTest",
            "dev.erst.gridgrind.engine.validation.ValidationComparisonOperatorsTest",
            "dev.erst.gridgrind.engine.backend.xssf.ExcelDataValidation*Test",
            "dev.erst.gridgrind.engine.backend.xssf.ExcelStructureValidationRewriteCoverageTest",
            "dev.erst.gridgrind.engine.backend.xssf.ExcelResidualValidationCoverageTest",
        ),
    )
}

tasks.named("pitest") {
    mustRunAfter(":core:pitest")
}

// Engine behavior is exercised through both runtime adapters; coverage evidence must include both.
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
    constraints {
        for (configuration in listOf("implementation", "testFixturesImplementation")) {
            add(configuration, libs.bnd.annotations)
            add(configuration, libs.osgi.resource)
            add(configuration, libs.commons.codec)
            add(configuration, libs.errorprone.annotations)
            add(configuration, libs.jspecify)
            add(configuration, libs.commons.lang3)
            add(configuration, libs.woodstox.core)
            add(configuration, libs.stax2.api)
            add(configuration, libs.slf4j.api)
        }
    }
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
    testImplementation(testFixtures(project(":core")))
    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}

tasks.named<JacocoReport>("jacocoTestReport") {
    dependsOn(tasks.named<Test>("test"))
    executionData.from(localCoverageExecutionData())
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
    testImplementation(libs.jackson.databind)
    add(paritySourceSet.implementationConfigurationName, libs.jackson.databind)
    add(paritySourceSet.implementationConfigurationName, libs.bcpkix)
    add(paritySourceSet.implementationConfigurationName, libs.bcprov)
    add(paritySourceSet.implementationConfigurationName, libs.bcutil)
    add(paritySourceSet.implementationConfigurationName, libs.xmlsec)
    add(paritySourceSet.runtimeOnlyConfigurationName, libs.slf4j.api)
}
val parityTest = tasks.register<Test>("parityTest") {
    description = "Runs the XSSF behavioural parity corpus."
    group = "verification"
    testClassesDirs = paritySourceSet.output.classesDirs
    // Parity invokes native owners directly, so its explicit runtime view includes all repairs.
    classpath = nativeOwnerClasses + paritySourceSet.runtimeClasspath
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
