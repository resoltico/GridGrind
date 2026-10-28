package dev.erst.gridgrind.buildlogic

import java.io.OutputStream
import org.gradle.api.GradleException
import org.gradle.api.file.DirectoryProperty
import org.gradle.api.file.ConfigurableFileCollection
import org.gradle.api.tasks.Classpath
import org.gradle.api.provider.Property
import org.gradle.api.tasks.Input
import org.gradle.api.tasks.InputDirectory
import org.gradle.api.tasks.JavaExec
import org.gradle.api.tasks.LocalState
import org.gradle.api.tasks.PathSensitive
import org.gradle.api.tasks.PathSensitivity
import org.gradle.api.tasks.TaskAction

/** Runs the standalone driver while owning the mutable corpus, findings and scratch lifetime. */
abstract class FuzzTask : JavaExec() {
    @get:Classpath abstract val instrumentationAgent: ConfigurableFileCollection
    @get:Input abstract val harness: Property<String>
    @get:Input abstract val seconds: Property<Int>
    @get:Input abstract val runs: Property<Int>
    @get:Input abstract val maxInputBytes: Property<Int>
    @get:InputDirectory @get:PathSensitive(PathSensitivity.RELATIVE)
    abstract val seeds: DirectoryProperty
    @get:LocalState abstract val workspace: DirectoryProperty

    @TaskAction
    override fun exec() {
        require(seconds.get() > 0) { "fuzzSeconds must be positive" }
        FuzzRunWorkspace.run(workspace.get().asFile.toPath(), seeds.get().asFile.toPath(), harness.get(), runs.get(), maxInputBytes.get()) { files ->
            val toolJvm = javaLauncher.get().metadata
            logger.lifecycle("Active JQF JVM: ${toolJvm.javaRuntimeVersion} (${toolJvm.vendor})")
            logger.lifecycle("Corpus payload ceiling: ${files.corpusByteCeiling} bytes; scratch, filesystem overhead and JVM are separate")
            systemProperty("java.io.tmpdir", files.temporary.toAbsolutePath().toString())
            systemProperty("gridgrind.fuzz.instrumentation", instrumentationAgent.singleFile.absolutePath)
            systemProperty("janala.includes", "dev.erst.gridgrind.core.,dev.erst.gridgrind.json.,dev.erst.gridgrind.engine.")
            systemProperty("janala.excludes", "")
            systemProperty("janala.verbose", "true")
            val instrumentationOutput = InstrumentationOutput(standardOutput ?: System.out)
            standardOutput = instrumentationOutput
            workingDir(files.temporary.toFile())
            jvmArgs("-javaagent:" + instrumentationAgent.singleFile.absolutePath)
            setArgs(listOf(
                harness.get(), files.corpus.toAbsolutePath().toString(),
                files.corpus.resolve("results").toAbsolutePath().toString(),
                files.findings.toAbsolutePath().toString(), seconds.get().toString(),
                runs.get().toString(), maxInputBytes.get().toString(),
            ))
            super.exec()
            if (instrumentationOutput.failed) {
                throw GradleException("JQF could not instrument a selected product class; inspect the instrumentation warning above")
            }
        }
    }
}

/** Detects Janala's fail-open warning without retaining an unbounded process log. */
internal class InstrumentationOutput(private val delegate: OutputStream) : OutputStream() {
    private val marker = "[WARNING] Could not instrument ".toByteArray(Charsets.UTF_8)
    private var matched = 0
    var failed = false
        private set

    override fun write(value: Int) {
        delegate.write(value)
        observe(value.toByte())
    }

    private fun observe(byte: Byte) {
        matched = if (byte == marker[matched]) matched + 1 else if (byte == marker[0]) 1 else 0
        if (matched == marker.size) {
            failed = true
            matched = 0
        }
    }

    override fun write(bytes: ByteArray, offset: Int, length: Int) {
        delegate.write(bytes, offset, length)
        for (index in offset until offset + length) observe(bytes[index])
    }

    override fun flush() = delegate.flush()
}
