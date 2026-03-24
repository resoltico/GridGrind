plugins {
    `java-library`
    id("gridgrind.java-conventions")
}

description = "Core GridGrind workbook automation engine"

dependencies {
    implementation(libs.poi.ooxml)
    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}
