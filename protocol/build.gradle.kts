plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "Agent-facing GridGrind protocol and execution layer"

dependencies {
    api(project(":engine"))
    implementation(libs.jackson.databind)
    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}
