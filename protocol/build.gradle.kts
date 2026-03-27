plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "Structured GridGrind protocol and request execution layer"

dependencies {
    api(project(":engine"))
    api(libs.jackson.databind)
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}
