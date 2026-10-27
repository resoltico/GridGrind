package dev.erst.gridgrind.buildlogic

import java.nio.channels.FileChannel
import java.nio.channels.OverlappingFileLockException
import java.nio.file.Files
import java.nio.file.LinkOption
import java.nio.file.Path
import java.nio.file.StandardOpenOption
import java.util.Comparator
import java.util.function.Consumer

/** Owns one bounded mutable corpus and temporary directory for an exclusive native fuzz run. */
object FuzzRunWorkspace {
    data class RunFiles(
        val corpus: Path,
        val temporary: Path,
        val findings: Path,
        val corpusByteCeiling: Long,
    )

    @JvmStatic
    fun run(
        root: Path,
        seeds: Path,
        harness: String,
        runs: Int,
        maxInputBytes: Int,
        action: Consumer<RunFiles>,
    ) {
        require(harness in setOf("protocol-request", "protocol-workflow", "engine-command-sequence", "xlsx-roundtrip")) {
            "Select -Pharness=protocol-request|protocol-workflow|engine-command-sequence|xlsx-roundtrip"
        }
        require(maxInputBytes > 0) { "fuzzMaxInputBytes must be positive" }
        val inputs = Files.list(seeds).use { paths -> paths.filter(Files::isRegularFile).sorted().toList() }
        require(inputs.isNotEmpty()) { "Selected fuzz harness has no committed seeds" }
        require(runs > inputs.size) { "fuzzRuns must exceed the committed seed count (${inputs.size})" }
        inputs.forEach { input ->
            require(Files.size(input) <= maxInputBytes) {
                "fuzzMaxInputBytes=$maxInputBytes would exclude committed seed $input (${Files.size(input)} bytes)"
            }
        }
        Files.createDirectories(root)
        FileChannel.open(root.resolve("run.lock"), StandardOpenOption.CREATE, StandardOpenOption.WRITE).use { channel ->
            val acquired = try {
                channel.tryLock()
            } catch (_: OverlappingFileLockException) {
                null
            }
            checkNotNull(acquired) { "Another fuzz run owns $root; use an isolated checkout for parallel runs" }.use {
                val findingsRoot = root.resolve("findings")
                val pendingFindings = Files.exists(findingsRoot) && Files.walk(findingsRoot).use { paths -> paths.anyMatch(Files::isRegularFile) }
                check(!pendingFindings) {
                    "Review and move existing fuzz findings from $findingsRoot before starting another run"
                }
                val corpusRoot = root.resolve("corpus")
                val temporary = root.resolve("tmp")
                deleteTree(corpusRoot)
                deleteTree(temporary)
                val corpus = Files.createDirectories(corpusRoot.resolve(harness))
                Files.createDirectories(temporary)
                val findings = Files.createDirectories(findingsRoot.resolve(harness))
                for (input in inputs) Files.copy(input, corpus.resolve(input.fileName))
                val ceiling = inputs.sumOf(Files::size) + runs.toLong() * maxInputBytes
                action.accept(RunFiles(corpus, temporary, findings, ceiling))
                // A failing/killed run keeps its scratch for diagnosis; the next admitted run resets it.
                deleteTree(temporary)
            }
        }
    }

    private fun deleteTree(root: Path) {
        if (!Files.exists(root, LinkOption.NOFOLLOW_LINKS)) return
        Files.walk(root).use { paths -> paths.sorted(Comparator.reverseOrder()).forEach(Files::delete) }
    }
}
