package dev.erst.gridgrind.buildlogic

import java.nio.file.Files
import org.gradle.api.Project

/** Retains the current configuration report and two preceding input comparisons after check. */
fun Project.registerGridGrindReportRetention() =
    tasks.register("pruneConfigurationCacheReports") { task ->
        task.description = "Retains the three newest project configuration-cache reports after check."
        val reportRoot = layout.buildDirectory.dir("reports/configuration-cache").get().asFile
        task.doLast {
            // Compute at execution: a Provider<FileCollection> is frozen by configuration caching.
            reportRoot.walkTopDown()
                .onEnter { !Files.isSymbolicLink(it.toPath()) }
                .filter { it.isFile && !Files.isSymbolicLink(it.toPath()) && it.name == "configuration-cache-report.html" }
                .sortedWith(compareByDescending<java.io.File> { it.lastModified() }.thenBy { it.path })
                .drop(3)
                .forEach { Files.delete(it.toPath()) }
            reportRoot.walkBottomUp()
                .onEnter { !Files.isSymbolicLink(it.toPath()) }
                .filter { it.isDirectory && it != reportRoot }
                .forEach { directory -> if (directory.list()?.isEmpty() == true) directory.delete() }
        }
    }
