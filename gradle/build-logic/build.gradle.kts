import org.gradle.api.tasks.compile.JavaCompile
import java.io.File
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
    }
}

// GridGrind's runtime, product modules, and shared build logic all target Java 26.
kotlin {
    jvmToolchain(26)
}

tasks.withType<KotlinCompile>().configureEach {
    incremental = false
    compilerOptions.jvmTarget.set(JvmTarget.JVM_26)
    doFirst {
        cleanDirectoryContents(destinationDirectory.get().asFile)
    }
}

tasks.withType<JavaCompile>().configureEach {
    options.release = 26
    doFirst {
        cleanDirectoryContents(destinationDirectory.get().asFile)
    }
}

fun cleanDirectoryContents(directory: File) {
    if (!directory.exists()) {
        directory.mkdirs()
        return
    }
    directory.listFiles()?.forEach(File::deleteRecursively)
}
