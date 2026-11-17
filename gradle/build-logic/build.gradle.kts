import org.gradle.api.tasks.compile.JavaCompile
import org.jetbrains.kotlin.gradle.dsl.JvmTarget
import org.jetbrains.kotlin.gradle.tasks.KotlinCompile
import java.util.Properties

plugins {
    `kotlin-dsl`
    `java-gradle-plugin`
}

dependencyLocking { lockAllConfigurations() }

repositories {
    mavenCentral()
    gradlePluginPortal()
}

dependencies {
    constraints {
        implementation(libs.jspecify)
        implementation(libs.errorprone.annotations)
        implementation(libs.slf4j.api)
        implementation(libs.guava)
        implementation(libs.commons.codec)
        implementation(libs.java.diff.utils)
        implementation(libs.jgit)
        implementation("org.jetbrains.kotlin:kotlin-stdlib-jdk7:$embeddedKotlinVersion")
        implementation("org.jetbrains.kotlin:kotlin-stdlib-jdk8:$embeddedKotlinVersion")
        implementation(libs.xz)
    }
    implementation(libs.asm)
    implementation(gradleApi())
    implementation(gradleKotlinDsl())
    implementation("com.diffplug.spotless:spotless-plugin-gradle:${libs.versions.spotless.get()}")
    implementation(
        "net.ltgt.gradle:gradle-errorprone-plugin:${libs.versions.errorprone.plugin.get()}",
    )
    implementation(
        "info.solidsoft.gradle.pitest:gradle-pitest-plugin:${libs.versions.pitest.gradle.plugin.get()}",
    )

    testImplementation(platform(libs.junit.bom))
    testImplementation(gradleTestKit())
    testImplementation(libs.jqf.instrument)
    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)
}

gradlePlugin {
    plugins {
        register("gridgrindNativeProvider") {
            id = "gridgrind.native-provider"
            implementationClass = "dev.erst.gridgrind.buildlogic.NativeProviderConventionsPlugin"
        }
        register("gridgrindRootConventions") {
            id = "gridgrind.root-conventions"
            implementationClass = "dev.erst.gridgrind.buildlogic.GridGrindRootConventionsPlugin"
        }
        register("gridgrindJavaConventions") {
            id = "gridgrind.java-conventions"
            implementationClass = "dev.erst.gridgrind.buildlogic.GridGrindJavaConventionsPlugin"
        }
        register("gridgrindFuzzConventions") {
            id = "gridgrind.fuzz-conventions"
            implementationClass = "dev.erst.gridgrind.buildlogic.GridGrindFuzzConventionsPlugin"
        }
        register("gridgrindMutationConventions") {
            id = "gridgrind.mutation-conventions"
            implementationClass = "dev.erst.gridgrind.buildlogic.GridGrindMutationConventionsPlugin"
        }
    }
}

val javaPins =
    providers.fileContents(layout.projectDirectory.file("../../gradle.properties")).asText.map { text ->
        Properties().apply { text.reader().use(::load) }
    }.get()
val javaVersion = javaPins.getProperty("gridgrindJavaVersion")?.toIntOrNull()
    ?: error("Root gradle.properties must declare an integer gridgrindJavaVersion")
val javaTargetVersion = javaPins.getProperty("gridgrindJavaTargetVersion")?.toIntOrNull()
    ?: error("Root gradle.properties must declare an integer gridgrindJavaTargetVersion")

// The included build derives its compiler target from the same pin as the product modules.
kotlin {
    jvmToolchain(javaVersion)
}

tasks.named<KotlinCompile>("compileKotlin") {
    compilerOptions.jvmTarget.set(JvmTarget.fromTarget(javaTargetVersion.toString()))
}

tasks.named<KotlinCompile>("compileTestKotlin") {
    compilerOptions.jvmTarget.set(JvmTarget.fromTarget(javaTargetVersion.toString()))
}

tasks.withType<JavaCompile>().configureEach {
    options.release = javaTargetVersion
}

tasks.test {
    useJUnitPlatform()
    systemProperty("gridgrind.repository.root", projectDir.parentFile.parentFile.absolutePath)
}
