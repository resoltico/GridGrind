import org.gradle.api.file.Directory
import org.gradle.api.provider.Provider
import org.gradle.api.tasks.compile.JavaCompile
import org.gradle.api.tasks.Delete
import org.jetbrains.kotlin.gradle.dsl.JvmTarget
import org.jetbrains.kotlin.gradle.tasks.KotlinCompile
import java.util.Properties

plugins {
    alias(libs.plugins.kotlin.jvm)
    `java-gradle-plugin`
}

dependencyLocking { lockAllConfigurations() }

repositories {
    mavenCentral()
    gradlePluginPortal()
}

dependencies {
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
    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)
}

gradlePlugin {
    plugins {
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

val javaRelease =
    providers.fileContents(layout.projectDirectory.file("../../gradle.properties")).asText.map { text ->
        val properties = Properties()
        text.reader().use(properties::load)
        properties.getProperty("gridgrindJavaVersion")?.toIntOrNull()
            ?: error("Root gradle.properties must declare an integer gridgrindJavaVersion")
    }.get()

// The included build derives its compiler target from the same pin as the product modules.
kotlin {
    jvmToolchain(javaRelease)
}

fun cleanDestinationTaskName(taskName: String): String =
    "clean${taskName.replaceFirstChar { if (it.isLowerCase()) it.titlecase() else it.toString() }}Destination"

fun registerCleanDestinationTask(
    taskName: String,
    destinationDirectory: Provider<Directory>,
) = tasks.register(cleanDestinationTaskName(taskName), Delete::class) { delete(destinationDirectory) }

val cleanCompileKotlinDestination =
    registerCleanDestinationTask(
        "compileKotlin",
        tasks.named<KotlinCompile>("compileKotlin").flatMap { it.destinationDirectory },
    )
val cleanCompileTestKotlinDestination =
    registerCleanDestinationTask(
        "compileTestKotlin",
        tasks.named<KotlinCompile>("compileTestKotlin").flatMap { it.destinationDirectory },
    )
val cleanCompileTestJavaDestination =
    registerCleanDestinationTask(
        "compileTestJava",
        tasks.named<JavaCompile>("compileTestJava").flatMap { it.destinationDirectory },
    )

tasks.named<KotlinCompile>("compileKotlin") {
    incremental = false
    compilerOptions.jvmTarget.set(JvmTarget.fromTarget(javaRelease.toString()))
    dependsOn(cleanCompileKotlinDestination)
}

tasks.named<KotlinCompile>("compileTestKotlin") {
    incremental = false
    compilerOptions.jvmTarget.set(JvmTarget.fromTarget(javaRelease.toString()))
    dependsOn(cleanCompileTestKotlinDestination)
}

tasks.named<JavaCompile>("compileTestJava") {
    options.release = javaRelease
    dependsOn(cleanCompileTestJavaDestination)
}

tasks.test {
    useJUnitPlatform()
    systemProperty("gridgrind.repository.root", projectDir.parentFile.parentFile.absolutePath)
}
