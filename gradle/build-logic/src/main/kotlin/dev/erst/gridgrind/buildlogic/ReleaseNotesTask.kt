package dev.erst.gridgrind.buildlogic

import java.time.LocalDate
import java.time.format.DateTimeParseException
import org.gradle.api.DefaultTask
import org.gradle.api.GradleException
import org.gradle.api.file.RegularFileProperty
import org.gradle.api.provider.Property
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.InputFile
import org.gradle.api.tasks.OutputFile
import org.gradle.api.tasks.PathSensitive
import org.gradle.api.tasks.PathSensitivity
import org.gradle.api.tasks.TaskAction

/** Prepares versioned release notes without creating tags or contacting a publication service. */
abstract class ReleaseNotesTask : DefaultTask() {
    @get:Input abstract val releaseVersion: Property<String>
    @get:InputFile @get:PathSensitive(PathSensitivity.NONE) abstract val changelog: RegularFileProperty
    @get:OutputFile abstract val notes: RegularFileProperty

    @TaskAction
    fun writeNotes() {
        val version = releaseVersion.get()
        if (!version.matches(Regex("[0-9]+\\.[0-9]+\\.[0-9]+"))) {
            throw GradleException("Release version must be X.Y.Z, derived from tag vX.Y.Z")
        }
        val lines = changelog.get().asFile.readLines()
        val header = "## [" + version + "]"
        val sections = lines.indices.filter { lines[it].startsWith(header) }
        if (sections.size != 1) throw GradleException("CHANGELOG.md must contain exactly one " + header + " section")
        val heading = lines[sections.single()]
        val datedHeading = Regex(Regex.escape(header) + " - ([0-9]{4}-[0-9]{2}-[0-9]{2})").matchEntire(heading)
            ?: throw GradleException("CHANGELOG.md release heading must be " + header + " - YYYY-MM-DD")
        try {
            LocalDate.parse(datedHeading.groupValues[1])
        } catch (exception: DateTimeParseException) {
            throw GradleException("CHANGELOG.md release heading has an invalid calendar date: " + heading, exception)
        }
        val body = lines.drop(sections.single() + 1).takeWhile { !it.startsWith("## ") }.joinToString("\n").trim()
        if (body.isEmpty()) throw GradleException("CHANGELOG.md release section is empty: " + version)
        val output = notes.get().asFile
        output.parentFile.mkdirs()
        output.writeText(body + "\n")
    }
}
