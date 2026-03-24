/**
 * Convention plugin applied to every GridGrind Java subproject.
 *
 * Centralises:
 *  - Java toolchain (version read once from gradle.properties)
 *  - sources JAR production
 *  - JUnit Platform test runner
 *  - JaCoCo report + coverage-verification tasks (100 % line, 95 % branch)
 */
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.tasks.bundling.Jar
import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.tasks.JacocoCoverageVerification
import org.gradle.testing.jacoco.tasks.JacocoReport

plugins {
    `java-base`
    jacoco
}

val gridgrindJavaVersion: Int =
    providers.gradleProperty("gridgrindJavaVersion").map(String::toInt).get()

extensions.configure<JavaPluginExtension> {
    toolchain {
        languageVersion = JavaLanguageVersion.of(gridgrindJavaVersion)
    }
    withSourcesJar()
}

tasks.withType<Jar>().configureEach {
    manifest {
        attributes(
            "Implementation-Title" to project.name,
            "Implementation-Version" to project.version,
            "Implementation-Vendor" to "Ervins Strauhmanis",
            "Implementation-License" to "MIT",
        )
    }
}

tasks.withType<Test>().configureEach {
    useJUnitPlatform()
}

tasks.named<JacocoReport>("jacocoTestReport") {
    dependsOn(tasks.withType<Test>())
    reports {
        xml.required = true
        html.required = true
    }
}

tasks.named<JacocoCoverageVerification>("jacocoTestCoverageVerification") {
    dependsOn(tasks.withType<Test>())
    violationRules {
        rule {
            limit {
                counter = "LINE"
                value = "COVEREDRATIO"
                minimum = "1.0".toBigDecimal()
            }
            limit {
                counter = "BRANCH"
                value = "COVEREDRATIO"
                minimum = "0.95".toBigDecimal()
            }
        }
    }
}
