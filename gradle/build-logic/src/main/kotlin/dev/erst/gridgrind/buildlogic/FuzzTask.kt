package dev.erst.gridgrind.buildlogic

import org.gradle.api.file.DirectoryProperty
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
            logger.lifecycle("Fuzz corpus ceiling: ${files.corpusByteCeiling} bytes; temporary workbooks, JVM and caches are separate")
            systemProperty("gridgrind.fuzz.harness", harness.get())
            systemProperty("java.io.tmpdir", files.temporary.toAbsolutePath().toString())
            workingDir(files.temporary.toFile())
            setArgs(listOf(
                "--target_class=dev.erst.gridgrind.fuzz.FuzzTarget",
                "--instrumentation_includes=dev.erst.gridgrind.core.**:dev.erst.gridgrind.json.**:dev.erst.gridgrind.engine.**",
                "--keep_going=1",
                "--reproducer_path=" + files.findings.toAbsolutePath(),
                "-max_total_time=" + seconds.get(),
                "-timeout=10",
                "-runs=" + runs.get(),
                "-max_len=" + maxInputBytes.get(),
                "-artifact_prefix=" + files.findings.toAbsolutePath() + "/",
                files.corpus.toAbsolutePath().toString(),
            ))
            super.exec()
        }
    }
}
