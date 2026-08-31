package dev.erst.gridgrind.buildlogic

import org.gradle.api.Project
import org.gradle.api.Task
import org.gradle.api.tasks.TaskProvider
import org.gradle.kotlin.dsl.register

internal fun Project.registerGridGrindMutationCheck(): TaskProvider<Task> =
    tasks.register("mutationCheck") { mutationTask ->
        mutationTask.group = "verification"
        mutationTask.description = "Runs the reviewed critical mutation-testing scopes."
        mutationTask.dependsOn(":contract:verifyPitestReport")
        mutationTask.dependsOn(":engine:verifyPitestReport")
        mutationTask.dependsOn(":cli:verifyPitestReport")
        mutationTask.dependsOn(":executor:verifyPitestReport")
    }
