package dev.erst.gridgrind.buildlogic

import java.io.File
import org.gradle.api.Project

internal data class GridGrindRepositoryLayout(
    val repositoryRoot: File,
    val productionPmdRuleset: File,
    val testPmdRuleset: File,
) {
    companion object {
        fun locate(project: Project): GridGrindRepositoryLayout {
            val repositoryRoot =
                generateSequence(project.projectDir) { current -> current.parentFile }
                    .firstOrNull(::isRepositoryRoot)
                    ?: throw IllegalStateException(
                        "Unable to locate the GridGrind repository root from ${project.projectDir}",
                    )
            val productionPmdRuleset = requiredFile(repositoryRoot, "gradle/pmd/ruleset.xml")
            val testPmdRuleset = requiredFile(repositoryRoot, "gradle/pmd/test-ruleset.xml")
            return GridGrindRepositoryLayout(
                repositoryRoot = repositoryRoot,
                productionPmdRuleset = productionPmdRuleset,
                testPmdRuleset = testPmdRuleset,
            )
        }

        private fun isRepositoryRoot(candidate: File): Boolean =
            candidate.resolve("settings.gradle.kts").isFile &&
                candidate.resolve("gradle/libs.versions.toml").isFile

        private fun requiredFile(repositoryRoot: File, relativePath: String): File =
            repositoryRoot.resolve(relativePath).also { candidate ->
                require(candidate.isFile) {
                    "Expected GridGrind build resource '$relativePath' under ${repositoryRoot.absolutePath}"
                }
            }
    }
}
