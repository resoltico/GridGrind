package dev.erst.gridgrind.buildlogic

import org.gradle.api.Project
import org.gradle.api.Task
import org.gradle.api.tasks.TaskProvider
import org.gradle.kotlin.dsl.register

internal fun Project.registerGridGrindMutationReport(): TaskProvider<Task> =
    tasks.register("mutationReport") { reportTask ->
        reportTask.group = "verification"
        reportTask.description = "Generates advisory mutation reports for the model and engine."
        reportTask.dependsOn(":contract:pitest")
        reportTask.dependsOn(":engine:pitest")
    }
