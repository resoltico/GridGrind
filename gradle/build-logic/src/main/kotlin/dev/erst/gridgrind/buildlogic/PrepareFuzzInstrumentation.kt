package dev.erst.gridgrind.buildlogic

import java.nio.file.Files
import java.nio.file.StandardCopyOption
import java.security.MessageDigest
import java.time.LocalDateTime
import java.util.zip.ZipEntry
import java.util.zip.ZipFile
import java.util.zip.ZipOutputStream
import org.gradle.api.DefaultTask
import org.gradle.api.GradleException
import org.gradle.api.file.RegularFileProperty
import org.gradle.api.provider.Property
import org.gradle.api.tasks.CacheableTask
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.InputFile
import org.gradle.api.tasks.OutputFile
import org.gradle.api.tasks.PathSensitive
import org.gradle.api.tasks.PathSensitivity
import org.gradle.api.tasks.TaskAction
import org.objectweb.asm.ClassReader
import org.objectweb.asm.ClassVisitor
import org.objectweb.asm.ClassWriter
import org.objectweb.asm.MethodVisitor
import org.objectweb.asm.Opcodes

/** Upgrades two pinned JQF visitor APIs while preserving the instrumenter's transformations. */
@CacheableTask
abstract class PrepareFuzzInstrumentation : DefaultTask() {
    @get:InputFile @get:PathSensitive(PathSensitivity.NONE)
    abstract val upstreamJar: RegularFileProperty
    @get:OutputFile abstract val agentJar: RegularFileProperty
    @get:Input abstract val legalNotice: Property<String>

    init {
        legalNotice.convention(
            checkNotNull(javaClass.getResourceAsStream("/licenses/jqf-instrument-3.0.txt"))
                .bufferedReader(Charsets.UTF_8).use { it.readText() },
        )
    }

    @TaskAction
    fun prepare() {
        val source = upstreamJar.get().asFile
        val digest = MessageDigest.getInstance("SHA-256")
        source.inputStream().buffered().use { input ->
            val buffer = ByteArray(8192)
            while (true) {
                val count = input.read(buffer)
                if (count < 0) break
                digest.update(buffer, 0, count)
            }
        }
        val actualHash = digest.digest().joinToString("") { "%02x".format(it) }
        check(actualHash == UPSTREAM_SHA256) {
            "Unreviewed JQF instrumentation artifact ($actualHash); review its visitor APIs before updating this patch"
        }
        val destination = agentJar.get().asFile.toPath()
        Files.createDirectories(destination.parent)
        val temporary = Files.createTempFile(destination.parent, "jqf-instrument-", ".jar")
        try {
            ZipFile(source).use { upstream ->
                val present = upstream.entries().asSequence().map { it.name }.toSet()
                check(present.containsAll(CONSTRUCTORS.keys)) { "Pinned JQF visitor classes are missing" }
                ZipOutputStream(Files.newOutputStream(temporary)).use { output ->
                    for (entry in upstream.entries().asSequence().sortedBy { it.name }) {
                        val bytes = upstream.getInputStream(entry).use { it.readBytes() }
                        val patched = CONSTRUCTORS[entry.name]?.let { rewriteVisitor(bytes, it) } ?: bytes
                        write(output, entry.name, patched)
                    }
                    write(output, "META-INF/LICENSE-JQF.txt", legalNotice.get().toByteArray(Charsets.UTF_8))
                    write(output, "META-INF/GRIDGRIND-INSTRUMENTATION.txt", PROVENANCE.toByteArray(Charsets.UTF_8))
                }
            }
            Files.move(temporary, destination, StandardCopyOption.REPLACE_EXISTING)
        } finally {
            Files.deleteIfExists(temporary)
        }
    }

    private fun rewriteVisitor(bytes: ByteArray, constructor: String): ByteArray {
        val reader = ClassReader(bytes)
        val writer = ClassWriter(reader, 0)
        var constructors = 0
        var changes = 0
        reader.accept(object : ClassVisitor(Opcodes.ASM9, writer) {
            override fun visitMethod(access: Int, name: String, descriptor: String, signature: String?, exceptions: Array<out String>?): MethodVisitor? {
                val method = super.visitMethod(access, name, descriptor, signature, exceptions)
                if (name != "<init>" || descriptor != constructor) return method
                constructors++
                return object : MethodVisitor(Opcodes.ASM9, method) {
                    override fun visitLdcInsn(value: Any?) {
                        if (value == Opcodes.ASM8) {
                            changes++
                            super.visitLdcInsn(Opcodes.ASM9)
                        } else {
                            super.visitLdcInsn(value)
                        }
                    }
                }
            }
        }, 0)
        if (constructors != 1 || changes != 1) {
            throw GradleException("Unexpected JQF visitor constructor in ${reader.className}: constructors=$constructors, API changes=$changes")
        }
        return writer.toByteArray()
    }

    private fun write(output: ZipOutputStream, name: String, bytes: ByteArray) {
        val entry = ZipEntry(name)
        entry.setTimeLocal(LocalDateTime.of(1980, 1, 1, 0, 0))
        output.putNextEntry(entry)
        output.write(bytes)
        output.closeEntry()
    }

    companion object {
        private const val UPSTREAM_SHA256 = "0d97c17f4a1b7f334ec26b0d9acd9f15457809eef7a16ecd1135adc5a985ad72"
        private val CONSTRUCTORS = mapOf(
            "janala/instrument/SnoopInstructionClassAdapter.class" to "(Lorg/objectweb/asm/ClassVisitor;Ljava/lang/String;)V",
            "janala/instrument/SnoopInstructionMethodAdapter.class" to "(Lorg/objectweb/asm/MethodVisitor;Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Ljava/lang/String;Ljanala/instrument/GlobalStateForInstrumentation;Z)V",
        )
        private val PROVENANCE = """
            GridGrind development-only JQF instrumentation preparation.
            Original: edu.berkeley.cs.jqf:jqf-instrument:3.0
            SHA-256: $UPSTREAM_SHA256
            Source: https://github.com/rohanpadhye/JQF/tree/2fd9c8dc7393276825f95e38c163f592df7e5533
            Modification: only the ASM8 API constant in the exact constructors of
            SnoopInstructionClassAdapter and SnoopInstructionMethodAdapter becomes ASM9.
            All remaining upstream entry bytes, including the agent manifest, are retained.
            ZIP timestamps are normalized; original BSD notices are included in LICENSE-JQF.txt.
            This tool is not included in GridGrind product archives or runtime images.
        """.trimIndent() + "\n"
    }
}
