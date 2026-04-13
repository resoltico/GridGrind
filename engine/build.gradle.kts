plugins {
    `java-library`
    `java-test-fixtures`
    id("gridgrind.java-conventions")
}

description = "Core GridGrind workbook automation engine"

dependencies {
    implementation(libs.poi.ooxml)
    implementation(libs.poi.ooxml.full)
    implementation(libs.xmlsec)
    testFixturesImplementation(libs.bcpkix)
    testFixturesImplementation(libs.bcprov)
    testFixturesImplementation(libs.bcutil)
    testFixturesImplementation(libs.poi.ooxml)
    testFixturesImplementation(libs.poi.ooxml.full)
    testFixturesImplementation(libs.xmlsec)
    testImplementation(testFixtures(project(":engine")))
    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.platform.launcher)
    testRuntimeOnly(libs.log4j.core)
}
