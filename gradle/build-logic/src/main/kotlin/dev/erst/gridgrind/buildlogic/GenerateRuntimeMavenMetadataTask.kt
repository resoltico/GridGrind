package dev.erst.gridgrind.buildlogic

import java.io.ByteArrayOutputStream
import java.io.File
import java.io.IOException
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.util.Properties
import java.util.zip.ZipFile
import javax.inject.Inject
import org.gradle.api.DefaultTask
import org.gradle.api.GradleException
import org.gradle.api.file.DirectoryProperty
import org.gradle.api.file.FileSystemOperations
import org.gradle.api.provider.ListProperty
import org.gradle.api.tasks.CacheableTask
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.InputFile
import org.gradle.api.tasks.Nested
import org.gradle.api.tasks.OutputDirectory
import org.gradle.api.tasks.PathSensitive
import org.gradle.api.tasks.PathSensitivity
import org.gradle.api.tasks.TaskAction

/** One actual resolved runtime artifact and the component identity assigned by its resolver. */
data class RuntimeMavenComponent(
    @get:Input val group: String,
    @get:Input val artifact: String,
    @get:Input val version: String,
    @get:InputFile @get:PathSensitive(PathSensitivity.NONE) val file: File,
)

/** Supplements missing standard component identities without replacing genuine vendor metadata. */
@CacheableTask
abstract class GenerateRuntimeMavenMetadataTask : DefaultTask() {
    @get:Nested abstract val components: ListProperty<RuntimeMavenComponent>
    @get:OutputDirectory abstract val outputDirectory: DirectoryProperty
    @get:Inject protected abstract val fileSystemOperations: FileSystemOperations

    @TaskAction
    @Throws(IOException::class)
    fun generateMetadata() {
        val identities = sortedMapOf<String, RuntimeMavenComponent>()
        val genuine = mutableSetOf<String>()
        for (component in components.get()) {
            requirePathSegment(component.group)
            requirePathSegment(component.artifact)
            if (component.version.isBlank()) throw GradleException("Runtime Maven version must not be blank")
            val path = "META-INF/maven/${component.group}/${component.artifact}/pom.properties"
            val prior = identities.putIfAbsent(path, component)
            if (prior != null && prior.version != component.version) {
                throw GradleException("Conflicting runtime Maven component versions at $path")
            }
            ZipFile(component.file).use { jar ->
                val entry = jar.getEntry(path)
                if (entry != null) {
                    val actual = Properties()
                    jar.getInputStream(entry).use(actual::load)
                    for ((key, value) in properties(component)) {
                        if (actual.getProperty(key.toString()) != value) {
                            throw GradleException("Upstream runtime Maven identity differs from resolution: $path ($key)")
                        }
                    }
                    genuine.add(path)
                }
            }
        }
        val root = outputDirectory.get().asFile.toPath()
        fileSystemOperations.delete { delete(root) }
        Files.createDirectories(root)
        for ((path, component) in identities) {
            if (path in genuine) continue
            // The standard writer supplies escaping; its optional timestamp comment is not identity.
            val encoded = ByteArrayOutputStream()
            properties(component).store(encoded, null)
            val text = encoded.toString(StandardCharsets.ISO_8859_1)
                .lineSequence().filter { it.isNotEmpty() && !it.startsWith("#") }.sorted()
                .joinToString("\n", postfix = "\n")
            val destination = root.resolve(path)
            Files.createDirectories(destination.parent)
            Files.writeString(destination, text, StandardCharsets.ISO_8859_1)
        }
    }

    private fun properties(component: RuntimeMavenComponent) = Properties().apply {
        setProperty("groupId", component.group)
        setProperty("artifactId", component.artifact)
        setProperty("version", component.version)
    }

    private fun requirePathSegment(value: String) {
        if (!value.matches(Regex("[A-Za-z0-9_.-]+")) || value == "." || value == "..") {
            throw GradleException("Runtime Maven group/artifact must be a Maven path segment: $value")
        }
    }
}
