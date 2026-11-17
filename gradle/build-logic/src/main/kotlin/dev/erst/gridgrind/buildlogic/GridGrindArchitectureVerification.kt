package dev.erst.gridgrind.buildlogic

import org.gradle.api.Project
import org.gradle.api.Task
import org.gradle.api.tasks.TaskProvider
import org.gradle.kotlin.dsl.register

internal fun Project.registerGridGrindArchitectureCheck(): TaskProvider<Task> =
    tasks.register("architectureCheck") { val architectureTask = this
        architectureTask.group = "verification"
        architectureTask.description =
            "Runs the mandatory bytecode-level product architecture rules."
        architectureTask.dependsOn(":architecture-tests:architectureTest")
        architectureTask.dependsOn(":architecture-tests:pmdArchitectureTest")
    }
