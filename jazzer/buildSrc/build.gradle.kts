import org.gradle.api.tasks.compile.JavaCompile

plugins {
    java
}

repositories {
    gradlePluginPortal()
}

dependencies {
    implementation(gradleApi())
}

// GridGrind's runtime and product-module baseline is Java 26. This nested Java-only build logic
// compiles with the Java 26 toolchain but still emits JVM 25 bytecode so it stays aligned with
// the current Kotlin build-logic ceiling without requiring a separate local Java 25 installation.
java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(26)
    }
}

tasks.withType<JavaCompile>().configureEach {
    options.release = 25
}
