plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "Java-first fluent authoring API over the canonical GridGrind contract"

dependencies {
    api(project(":executor"))
    testImplementation(libs.junit.jupiter)
    testImplementation(testFixtures(project(":engine")))
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}
