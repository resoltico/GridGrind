import org.gradle.api.file.Directory
import org.gradle.api.provider.Provider
import org.gradle.api.tasks.compile.JavaCompile
import org.gradle.api.tasks.Delete
import org.jetbrains.kotlin.gradle.dsl.JvmTarget
import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

plugins {
    alias(libs.plugins.kotlin.jvm)
    `java-gradle-plugin`
}

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
    implementation("net.sourceforge.pmd:pmd-java:${libs.versions.pmd.get()}")
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
        register("gridgrindJazzerConventions") {
            id = "gridgrind.jazzer-conventions"
            implementationClass = "dev.erst.gridgrind.buildlogic.GridGrindJazzerConventionsPlugin"
        }
        register("gridgrindMutationConventions") {
            id = "gridgrind.mutation-conventions"
            implementationClass = "dev.erst.gridgrind.buildlogic.GridGrindMutationConventionsPlugin"
        }
    }
}

// GridGrind's runtime, product modules, and shared build logic all target Java 26.
kotlin {
    jvmToolchain(26)
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
val cleanCompileJavaDestination =
    registerCleanDestinationTask(
        "compileJava",
        tasks.named<JavaCompile>("compileJava").flatMap { it.destinationDirectory },
    )
val cleanCompileTestJavaDestination =
    registerCleanDestinationTask(
        "compileTestJava",
        tasks.named<JavaCompile>("compileTestJava").flatMap { it.destinationDirectory },
    )

tasks.named<KotlinCompile>("compileKotlin") {
    incremental = false
    compilerOptions.jvmTarget.set(JvmTarget.JVM_26)
    dependsOn(cleanCompileKotlinDestination)
}

tasks.named<KotlinCompile>("compileTestKotlin") {
    incremental = false
    compilerOptions.jvmTarget.set(JvmTarget.JVM_26)
    dependsOn(cleanCompileTestKotlinDestination)
}

tasks.named<JavaCompile>("compileJava") {
    options.release = 26
    dependsOn(cleanCompileJavaDestination)
}

tasks.named<JavaCompile>("compileTestJava") {
    options.release = 26
    dependsOn(cleanCompileTestJavaDestination)
}

tasks.test {
    useJUnitPlatform()
    systemProperty("gridgrind.repository.root", projectDir.parentFile.parentFile.absolutePath)
}
