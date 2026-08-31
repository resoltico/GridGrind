package dev.erst.gridgrind.buildlogic

import org.gradle.api.Project
import org.gradle.api.Task
import org.gradle.api.tasks.TaskProvider
import org.gradle.kotlin.dsl.register

internal fun Project.registerGridGrindArchitectureCheck(): TaskProvider<Task> =
    tasks.register("architectureCheck") { architectureTask ->
        architectureTask.group = "verification"
        architectureTask.description =
            "Runs the mandatory bytecode-level product architecture rules."
        architectureTask.dependsOn(":executor:architectureTest")
        architectureTask.dependsOn(":executor:architectureTestContract")
        architectureTask.dependsOn(":executor:pmdArchitectureTest")
    }
