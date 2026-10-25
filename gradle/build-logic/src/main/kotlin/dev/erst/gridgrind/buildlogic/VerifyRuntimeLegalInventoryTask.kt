package dev.erst.gridgrind.buildlogic

import java.io.IOException
import java.nio.file.Files
import java.util.TreeSet
import org.gradle.api.DefaultTask
import org.gradle.api.GradleException
import org.gradle.api.file.ConfigurableFileCollection
import org.gradle.api.file.RegularFileProperty
import org.gradle.api.provider.SetProperty
import org.gradle.api.tasks.Classpath
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.InputFile
import org.gradle.api.tasks.PathSensitive
import org.gradle.api.tasks.PathSensitivity
import org.gradle.api.tasks.TaskAction

/** Fails packaging when its external runtime artifacts or audited NOTICE markers drift. */
abstract class VerifyRuntimeLegalInventoryTask : DefaultTask() {
    @get:Classpath abstract val runtimeArtifacts: ConfigurableFileCollection

    @get:Input abstract val auditedRuntimeArtifactNames: SetProperty<String>

    @get:Input abstract val auditedNoticeMarkers: SetProperty<String>

    @get:InputFile
    @get:PathSensitive(PathSensitivity.RELATIVE)
    abstract val noticeFile: RegularFileProperty

    @TaskAction
    @Throws(IOException::class)
    fun verifyInventory() {
        val resolvedArtifactNames = TreeSet<String>()
        runtimeArtifacts.files.forEach { artifact ->
            if (!resolvedArtifactNames.add(artifact.name)) {
                throw GradleException(
                    "Runtime legal inventory contains duplicate artifact name: ${artifact.name}",
                )
            }
        }
        val auditedArtifactNames = auditedRuntimeArtifactNames.get().toSortedSet()
        val unaudited = (resolvedArtifactNames - auditedArtifactNames).toSortedSet()
        val noLongerResolved = (auditedArtifactNames - resolvedArtifactNames).toSortedSet()
        if (unaudited.isNotEmpty() || noLongerResolved.isNotEmpty()) {
            throw GradleException(
                buildString {
                    append("Runtime legal inventory drifted.")
                    if (unaudited.isNotEmpty()) append(" Unaudited: ").append(unaudited).append('.')
                    if (noLongerResolved.isNotEmpty()) append(" No longer resolved: ").append(noLongerResolved).append('.')
                    append(" Audit the changed artifacts, then update NOTICE and the audited artifact set together.")
                },
            )
        }
        val noticeText = Files.readString(noticeFile.get().asFile.toPath())
        val missingMarkers = auditedNoticeMarkers.get().filterNot(noticeText::contains).toSortedSet()
        if (missingMarkers.isNotEmpty()) {
            throw GradleException("NOTICE is missing audited runtime markers: $missingMarkers")
        }
    }
}
