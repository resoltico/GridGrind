plugins {
    `kotlin-dsl`
}

repositories {
    gradlePluginPortal()
}

// Kotlin does not yet support JVM target 26. Build logic in buildSrc compiles to
// JVM 25 bytecode, which runs without issue on the Java 26 runtime used everywhere
// else. Production modules (engine, protocol, cli) are unaffected and still target 26.
kotlin {
    jvmToolchain(25)
}
