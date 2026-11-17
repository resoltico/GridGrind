package dev.erst.gridgrind.buildlogic

import groovy.json.JsonOutput
import groovy.json.JsonSlurper
import java.io.File
import java.nio.file.Files
import java.security.MessageDigest
import java.util.HexFormat
import java.util.Properties
import java.util.zip.ZipFile
import org.gradle.api.DefaultTask
import org.gradle.api.GradleException
import org.gradle.api.file.RegularFileProperty
import org.gradle.api.provider.ListProperty
import org.gradle.api.provider.Property
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.InputFile
import org.gradle.api.tasks.Nested
import org.gradle.api.tasks.OutputFile
import org.gradle.api.tasks.PathSensitive
import org.gradle.api.tasks.PathSensitivity
import org.gradle.api.tasks.TaskAction
import org.gradle.work.DisableCachingByDefault

/** Publishes resolver-owned runtime identities and a separate preparation-only artifact locator map. */
@DisableCachingByDefault(because = "The preparation-only file map records local resolver paths")
abstract class WriteRuntimeComponentInventoryTask : DefaultTask() {
    @get:Nested abstract val components: ListProperty<RuntimeMavenComponent>
    @get:Input abstract val legalFiles: ListProperty<String>
    @get:InputFile @get:PathSensitive(PathSensitivity.NONE) abstract val nativeCapsule: RegularFileProperty
    @get:Input abstract val nativeGroup: Property<String>
    @get:Input abstract val nativeArtifact: Property<String>
    @get:Input abstract val nativeVersion: Property<String>
    @get:OutputFile abstract val inventoryFile: RegularFileProperty
    @get:OutputFile abstract val resolvedFilesFile: RegularFileProperty

    // Path changes must refresh the local map even when identical artifact bytes are restored.
    @get:Input
    val resolvedLocations: List<String>
        get() = (components.get().map { it.file.absolutePath } + nativeCapsule.get().asFile.absolutePath).sorted()

    @TaskAction
    fun writeInventory() {
        val external = components.get()
        if (external.isEmpty()) throw GradleException("Resolved external runtime component inventory is empty")
        val own = RuntimeMavenComponent(nativeGroup.get(), nativeArtifact.get(), nativeVersion.get(), nativeCapsule.get().asFile)
        verifyNativeIdentity(own)
        val actual = (external.map { it to "external" } + (own to "project"))
            .sortedBy { (component, _) -> "${component.group}:${component.artifact}:${component.version}" }
        val identities = mutableSetOf<String>()
        val publicRows = mutableListOf<Map<String, String>>()
        val localRows = mutableListOf<Map<String, String>>()
        for ((component, origin) in actual) {
            val identity = "${component.group}:${component.artifact}:${component.version}"
            if (!identities.add(identity)) throw GradleException("Repeated resolved runtime component identity: $identity")
            val row = linkedMapOf("group" to component.group, "artifact" to component.artifact,
                "version" to component.version, "origin" to origin, "sha256" to sha256(component.file))
            publicRows.add(row)
            localRows.add(row + ("file" to component.file.absolutePath))
        }
        val legal = legalFiles.get()
        if (legal.isEmpty() || legal.toSet().size != legal.size)
            throw GradleException("Packaged runtime legal-file inventory is empty or repeated")
        writeJson(inventoryFile.get().asFile, mapOf("formatVersion" to 1, "components" to publicRows, "legalFiles" to legal))
        writeJson(resolvedFilesFile.get().asFile, mapOf("formatVersion" to 1, "components" to localRows, "legalFiles" to legal))
    }

    private fun verifyNativeIdentity(component: RuntimeMavenComponent) {
        ZipFile(component.file).use { capsule ->
            val manifestEntry = capsule.getEntry("META-INF/gridgrind/native-provider/manifest.json")
                ?: throw GradleException("Resolved native provider capsule has no component manifest")
            val manifest = capsule.getInputStream(manifestEntry).use { JsonSlurper().parse(it) as? Map<*, *> }
                ?: throw GradleException("Resolved native provider manifest is not an object")
            val identity = manifest["component"] as? Map<*, *>
                ?: throw GradleException("Resolved native provider manifest has no component identity")
            val expected = mapOf("group" to component.group, "artifact" to component.artifact, "version" to component.version)
            if (identity != expected) throw GradleException("Native provider manifest differs from its consumed project coordinate")
            val entry = capsule.getEntry("META-INF/maven/${component.group}/${component.artifact}/pom.properties")
                ?: throw GradleException("Resolved native provider has no own Maven component identity")
            val properties = Properties()
            capsule.getInputStream(entry).use(properties::load)
            if (properties.getProperty("groupId") != component.group ||
                properties.getProperty("artifactId") != component.artifact ||
                properties.getProperty("version") != component.version) {
                throw GradleException("Native provider Maven identity differs from its consumed project coordinate")
            }
        }
    }

    private fun sha256(file: File): String {
        val hash = MessageDigest.getInstance("SHA-256")
        file.inputStream().use { input ->
            val buffer = ByteArray(8192)
            while (true) {
                val count = input.read(buffer)
                if (count < 0) break
                hash.update(buffer, 0, count)
            }
        }
        return HexFormat.of().formatHex(hash.digest())
    }

    private fun writeJson(file: File, value: Any) {
        Files.createDirectories(file.toPath().parent)
        Files.writeString(file.toPath(), JsonOutput.prettyPrint(JsonOutput.toJson(value)) + "\n")
    }
}
