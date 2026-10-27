package dev.erst.gridgrind.buildlogic

import org.gradle.api.Project
import org.gradle.api.tasks.TaskCollection
import org.gradle.api.tasks.testing.Test
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension
import org.gradle.api.artifacts.VersionCatalog
import org.gradle.api.artifacts.VersionCatalogsExtension
import org.gradle.kotlin.dsl.getByType

internal fun Project.versionCatalog(name: String = "libs"): VersionCatalog =
    extensions.getByType<VersionCatalogsExtension>().named(name)

internal fun Project.configureGridGrindRepositories() {
    repositories.mavenCentral()
}

/** External opt-in suites disable JaCoCo and must not become ordinary coverage prerequisites. */
fun Project.coverageTestTasks(): TaskCollection<Test> =
    tasks.withType(Test::class.java).matching { test ->
        test.extensions.getByType(JacocoTaskExtension::class.java).isEnabled
    }
