package dev.erst.gridgrind.buildlogic

import java.io.IOException
import java.nio.file.Files
import org.gradle.api.DefaultTask
import org.gradle.api.file.DirectoryProperty
import org.gradle.api.provider.ListProperty
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.Internal
import org.gradle.api.tasks.TaskAction

/** Removes stale top-level JARs while preserving the declared archives and other outputs. */
abstract class PruneJarOutputsTask : DefaultTask() {
    @get:Input abstract val expectedArchiveFileNames: ListProperty<String>

    @get:Internal abstract val libsDirectory: DirectoryProperty

    @TaskAction
    @Throws(IOException::class)
    fun pruneJarOutputs() {
        val directory = libsDirectory.get().asFile.toPath()
        if (!Files.isDirectory(directory)) return
        val expected = java.util.Set.copyOf(expectedArchiveFileNames.getOrElse(emptyList()))
        Files.list(directory).use { children ->
            children
                .filter { Files.isRegularFile(it) }
                .filter { it.fileName.toString().endsWith(".jar") }
                .filter { it.fileName.toString() !in expected }
                .forEach { candidate ->
                    try {
                        Files.deleteIfExists(candidate)
                    } catch (exception: IOException) {
                        throw IllegalStateException("Failed pruning stale jar output $candidate", exception)
                    }
                }
        }
    }
}
