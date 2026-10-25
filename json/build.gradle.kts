plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "JSON codecs and tolerant request intake"

dependencies {
    api(project(":core"))
    implementation(libs.jackson.annotations)
    implementation(libs.jackson.databind)
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testImplementation(libs.poi.ooxml.full)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}

tasks.withType<Test>().configureEach {
    systemProperty("gridgrind.root.dir", rootProject.projectDir.absolutePath)
}
