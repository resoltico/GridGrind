pluginManagement {
    includeBuild("gradle/build-logic")
}

plugins {
    id("org.gradle.toolchains.foojay-resolver-convention") version "1.0.0"
}

rootProject.name = "GridGrind"
include("excel-foundation", "engine", "contract", "architecture-tests", "authoring-java", "cli")
