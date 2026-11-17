package dev.erst.gridgrind.buildlogic

import groovy.json.JsonOutput
import groovy.json.JsonSlurper
import java.io.File
import java.nio.charset.StandardCharsets
import java.security.MessageDigest
import java.util.zip.ZipFile
import javax.inject.Inject
import org.gradle.api.DefaultTask
import org.gradle.api.file.ConfigurableFileCollection
import org.gradle.api.file.DirectoryProperty
import org.gradle.api.file.FileSystemOperations
import org.gradle.api.file.RegularFileProperty
import org.gradle.api.provider.Property
import org.gradle.api.tasks.CacheableTask
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.InputDirectory
import org.gradle.api.tasks.InputFile
import org.gradle.api.tasks.InputFiles
import org.gradle.api.tasks.Internal
import org.gradle.api.tasks.OutputDirectory
import org.gradle.api.tasks.PathSensitive
import org.gradle.api.tasks.PathSensitivity
import org.gradle.api.tasks.TaskAction

internal object NativeRepairFiles {
    fun sha(file: File) = sha(file.readBytes())
    fun sha(bytes: ByteArray): String = MessageDigest.getInstance("SHA-256").digest(bytes).joinToString("") { "%02x".format(java.util.Locale.ROOT, it) }
    fun check(file: File, expected: String) { require(sha(file) == expected) { "Native repair input checksum differs: ${file.name}" } }
    @Suppress("UNCHECKED_CAST")
    fun manifest(file: File): Map<String, Any> = JsonSlurper().parse(file) as Map<String, Any>
    @Suppress("UNCHECKED_CAST")
    fun rows(value: Any?): List<Map<String, Any>> = value as List<Map<String, Any>>
    fun child(root: File, relative: String): File {
        val resolved = root.toPath().resolve(relative).normalize()
        require(!File(relative).isAbsolute && resolved.startsWith(root.toPath()) && resolved != root.toPath()) { "Native repair path escapes its owner: $relative" }
        return resolved.toFile()
    }
    fun json(file: File, value: Any) { file.parentFile.mkdirs(); file.writeText(JsonOutput.prettyPrint(JsonOutput.toJson(value)) + "\n", Charsets.UTF_8) }
    /** Applies only exact pinned unified hunks, without offsets, fuzz or a platform patch tool. */
    fun applyPatch(original: String, patch: String): String {
        val source = original.split('\n'); val changes = patch.split('\n'); val result = mutableListOf<String>()
        var cursor = 0; var index = 0; var hunks = 0
        val header = Regex("@@ -(\\d+)(?:,(\\d+))? \\+(\\d+)(?:,(\\d+))? @@.*")
        while (index < changes.size) {
            val line = changes[index++]
            if (line.isEmpty() || line.startsWith("--- ") || line.startsWith("+++ ")) continue
            val match = header.matchEntire(line) ?: error("Unsupported native patch header: $line")
            val start = match.groupValues[1].toInt() - 1
            val oldCount = match.groupValues[2].ifEmpty { "1" }.toInt()
            val newCount = match.groupValues[4].ifEmpty { "1" }.toInt()
            require(start >= cursor && start <= source.size) { "Native patch hunk order differs" }
            result.addAll(source.subList(cursor, start)); cursor = start
            var old = 0; var added = 0
            while (old < oldCount || added < newCount) {
                require(index < changes.size) { "Truncated native patch" }
                val change = changes[index++]; require(change.isNotEmpty()) { "Invalid native patch line" }
                val text = change.substring(1)
                when (change[0]) {
                    ' ', '-' -> { require(cursor < source.size && source[cursor] == text) { "Native patch context differs at source line ${cursor + 1}" }; cursor++; old++; if (change[0] == ' ') { result.add(text); added++ } }
                    '+' -> { result.add(text); added++ }
                    else -> error("Unsupported native patch line")
                }
                require(old <= oldCount && added <= newCount) { "Native patch hunk count differs" }
            }
            hunks++
        }
        require(hunks > 0) { "Native patch has no changes" }
        result.addAll(source.subList(cursor, source.size)); return result.joinToString("\n")
    }
}

/** Checks the reviewed source/patch pins and produces only selected compiler sources. */
@CacheableTask
abstract class PrepareNativeProviderSources : DefaultTask() {
    @get:Internal abstract val recipeDirectory: DirectoryProperty
    @get:InputFiles @get:PathSensitive(PathSensitivity.RELATIVE) abstract val recipeFiles: ConfigurableFileCollection
    @get:InputFile @get:PathSensitive(PathSensitivity.NONE) abstract val poiSources: RegularFileProperty
    @get:InputFile @get:PathSensitive(PathSensitivity.NONE) abstract val ooxmlSources: RegularFileProperty
    @get:Input abstract val compilerIdentity: Property<String>
    @get:Input abstract val requiredCompilerRuntime: Property<String>
    @get:OutputDirectory abstract val outputDirectory: DirectoryProperty
    @get:Inject protected abstract val fileSystemOperations: FileSystemOperations
    @TaskAction fun prepare() {
        require(compilerIdentity.get() == "Eclipse Temurin:${requiredCompilerRuntime.get()}") { "Native repairs require the pinned Temurin compiler; found ${compilerIdentity.get()}" }
        val recipes = recipeDirectory.get().asFile; val manifestFile = recipes.resolve("repairs.json")
        val manifest = NativeRepairFiles.manifest(manifestFile)
        require(manifest["formatVersion"] == 1) { "Unsupported native repair manifest" }
        val archives = mapOf("poi" to poiSources.get().asFile, "poi-ooxml" to ooxmlSources.get().asFile)
        val sourceArtifacts = NativeRepairFiles.rows(manifest["sourceArtifacts"])
        require(sourceArtifacts.map { it["id"] }.toSet() == archives.keys) { "Native source artifact roles differ" }
        sourceArtifacts.forEach { NativeRepairFiles.check(archives.getValue(it["id"] as String), it["sha256"] as String) }
        val root = outputDirectory.get().asFile
        fileSystemOperations.delete { delete(root) }; root.mkdirs()
        val rows = mutableListOf<Map<String, Any>>()
        for (repair in NativeRepairFiles.rows(manifest["repairs"])) {
            val component = repair["component"] as String; val path = repair["sourcePath"] as String
            val patch = NativeRepairFiles.child(recipes, repair["patchPath"] as String)
            NativeRepairFiles.check(patch, repair["patchSha256"] as String)
            val original = ZipFile(archives.getValue(component)).use { jar ->
                val entry = jar.getEntry(path) ?: error("Pinned native source missing: $path")
                jar.getInputStream(entry).use { it.readAllBytes() }
            }
            require(NativeRepairFiles.sha(original) == repair["upstreamSourceSha256"]) { "Pinned native source checksum differs: $path" }
            val changed = NativeRepairFiles.applyPatch(String(original, StandardCharsets.UTF_8), patch.readText(Charsets.UTF_8))
            val output = NativeRepairFiles.child(root.resolve(component), path)
            output.parentFile.mkdirs(); output.writeText(changed, Charsets.UTF_8)
            repair["modifiedSourceSha256"]?.let { NativeRepairFiles.check(output, it as String) }
            rows.add(repair + mapOf("modifiedSourceSha256" to NativeRepairFiles.sha(output), "license" to "Apache-2.0 AND AGPL-3.0-only"))
        }
        require(rows.isNotEmpty()) { "Native repair manifest is empty" }
        for (helper in NativeRepairFiles.rows(manifest["helpers"])) {
            val source = NativeRepairFiles.child(recipes, helper["sourcePath"] as String)
            NativeRepairFiles.check(source, helper["sha256"] as String)
            val relative = (helper["sourcePath"] as String).removePrefix("src/helper/java/")
            require(relative != helper["sourcePath"]) { "Unexpected native helper source role" }
            val output = NativeRepairFiles.child(root.resolve("helper"), relative)
            output.parentFile.mkdirs(); source.copyTo(output, overwrite = true)
            rows.add(helper + mapOf("component" to "helper", "compiledSourcePath" to relative))
        }
        NativeRepairFiles.json(root.resolve("source-provenance.json"), mapOf("formatVersion" to 1,
            "recipeManifestSha256" to NativeRepairFiles.sha(manifestFile), "sourceArtifacts" to sourceArtifacts, "sources" to rows))
    }
}

/** Creates the resource-only runtime payload and source/legal provenance from actual outputs. */
@CacheableTask
abstract class AssembleNativeProviderResources : DefaultTask() {
    @get:InputDirectory @get:PathSensitive(PathSensitivity.RELATIVE) abstract val preparedSources: DirectoryProperty
    @get:Internal abstract val recipeDirectory: DirectoryProperty
    @get:InputFiles @get:PathSensitive(PathSensitivity.RELATIVE) abstract val recipeFiles: ConfigurableFileCollection
    @get:InputFiles @get:PathSensitive(PathSensitivity.RELATIVE) abstract val classDirectories: ConfigurableFileCollection
    @get:InputFiles @get:PathSensitive(PathSensitivity.NONE) abstract val sourceArchives: ConfigurableFileCollection
    @get:InputFile @get:PathSensitive(PathSensitivity.NONE) abstract val apacheLicense: RegularFileProperty
    @get:InputFile @get:PathSensitive(PathSensitivity.NONE) abstract val projectLicense: RegularFileProperty
    @get:Input abstract val componentVersion: Property<String>
    @get:Input abstract val compilerIdentity: Property<String>
    @get:Input abstract val targetRelease: Property<Int>
    @get:OutputDirectory abstract val outputDirectory: DirectoryProperty
    @get:Inject protected abstract val fileSystemOperations: FileSystemOperations
    @TaskAction fun assemble() {
        val output = outputDirectory.get().asFile; fileSystemOperations.delete { delete(output) }
        val prefix = output.resolve("META-INF/gridgrind/native-provider"); prefix.mkdirs()
        val classes = sortedMapOf<String, String>()
        for (directory in classDirectories.files.sortedBy { it.name }) {
            for (file in directory.walkTopDown().filter { it.isFile && it.extension == "class" }.sortedBy { it.path }) {
                val relative = file.relativeTo(directory).invariantSeparatorsPath
                require(relative.startsWith("org/apache/poi/") && relative != "module-info.class") { "Unexpected native repair class: $relative" }
                val bytes = file.readBytes(); require(bytes.size > 8 && bytes[6].toInt() == 0 && (bytes[7].toInt() and 255) == targetRelease.get() + 44) { "Native repair bytecode target differs: $relative" }
                require(classes.put(relative, NativeRepairFiles.sha(bytes)) == null) { "Duplicate native repair class: $relative" }
                val destination = prefix.resolve("classes/$relative"); destination.parentFile.mkdirs(); file.copyTo(destination, overwrite = true)
            }
        }
        require(classes.isNotEmpty()) { "Native repair compiler output is empty" }
        val prepared = preparedSources.get().asFile
        val provenance = NativeRepairFiles.manifest(prepared.resolve("source-provenance.json"))
        for (row in NativeRepairFiles.rows(provenance["sources"])) {
            val relative = (row["compiledSourcePath"] ?: row["sourcePath"]) as String
            require(classes.containsKey(relative.removeSuffix(".java") + ".class")) { "Selected repair compiler output missing: $relative" }
        }
        prefix.resolve("required-classes.list").writeText(classes.keys.joinToString("\n", postfix = "\n") { it.removeSuffix(".class").replace('/', '.') }, Charsets.UTF_8)
        NativeRepairFiles.json(prefix.resolve("manifest.json"), provenance + mapOf("component" to mapOf("group" to "dev.erst.gridgrind", "artifact" to "native-provider", "version" to componentVersion.get()), "compiler" to compilerIdentity.get(), "targetRelease" to targetRelease.get(), "classes" to classes))
        for (archive in sourceArchives.files) { val dest = prefix.resolve("sources/upstream/${archive.name}"); dest.parentFile.mkdirs(); archive.copyTo(dest, overwrite = true) }
        for (role in listOf("poi", "poi-ooxml", "helper")) { prepared.resolve(role).copyRecursively(prefix.resolve("sources/modified/$role"), overwrite = true) }
        recipeDirectory.get().asFile.resolve("patches").copyRecursively(prefix.resolve("sources/patches"), overwrite = true)
        recipeDirectory.get().asFile.resolve("repairs.json").copyTo(prefix.resolve("repairs.json"), overwrite = true)
        apacheLicense.get().asFile.copyTo(prefix.resolve("LICENSE-APACHE-2.0"), overwrite = true)
        projectLicense.get().asFile.copyTo(prefix.resolve("LICENSE-AGPL-3.0-only"), overwrite = true)
        prefix.resolve("NOTICE").writeText("GridGrind native-provider ${componentVersion.get()}\nCopyright (c) 2026 Ervins Strauhmanis (GridGrind modifications and helper).\nReviewed repairs to Apache POI 6.0.0. Original upstream sources retain Apache-2.0 terms and headers; GridGrind modifications and helper are AGPL-3.0-only. Parent upstream class bytes remain unmodified.\n", Charsets.UTF_8)
        val maven = output.resolve("META-INF/maven/dev.erst.gridgrind/native-provider/pom.properties"); maven.parentFile.mkdirs()
        maven.writeText("artifactId=native-provider\ngroupId=dev.erst.gridgrind\nversion=${componentVersion.get()}\n", Charsets.UTF_8)
    }
}

/** Adds selected compiler directories to the existing vendor modules, never as split modules. */
abstract class NativeProviderModuleArguments : org.gradle.process.CommandLineArgumentProvider {
    @get:InputFiles @get:PathSensitive(PathSensitivity.RELATIVE) abstract val poiClasses: ConfigurableFileCollection
    @get:InputFiles @get:PathSensitive(PathSensitivity.RELATIVE) abstract val ooxmlClasses: ConfigurableFileCollection
    override fun asArguments(): Iterable<String> = listOf("--patch-module", "org.apache.poi.poi=${poiClasses.asPath}", "--patch-module", "org.apache.poi.ooxml=${ooxmlClasses.asPath}")
}
