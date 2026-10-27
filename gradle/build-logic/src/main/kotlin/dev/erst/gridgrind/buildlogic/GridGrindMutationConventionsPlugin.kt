package dev.erst.gridgrind.buildlogic

import info.solidsoft.gradle.pitest.PitestPlugin
import info.solidsoft.gradle.pitest.PitestPluginExtension
import org.gradle.api.GradleException
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.artifacts.VersionCatalogsExtension
import org.gradle.api.tasks.Delete
import org.gradle.api.tasks.SourceSetContainer
import org.gradle.jvm.toolchain.JavaLanguageVersion
import org.gradle.jvm.toolchain.JavaToolchainService
import java.math.BigDecimal
import org.gradle.kotlin.dsl.configure
import org.gradle.kotlin.dsl.getByType

/** Shared advisory PIT reporting for explicitly scoped product modules. */
class GridGrindMutationConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        project.pluginManager.apply(PitestPlugin::class.java)
        val versions = project.extensions.getByType<VersionCatalogsExtension>().named("libs")
        val javaVersion = project.providers.gradleProperty("gridgrindJavaVersion").map(String::toInt)
        val toolchains = project.extensions.getByType<JavaToolchainService>()
        val sourceSets = project.extensions.getByType<SourceSetContainer>()
        val pitestTestSourceSets = setOf(sourceSets.getByName("test"))
        val reportDirectory = project.layout.buildDirectory.dir("reports/pitest")
        val scratchDirectory = project.layout.buildDirectory.dir("tmp/pitest")
        val mutationWorkers = minOf(
            4,
            Runtime.getRuntime().availableProcessors(),
            project.gradle.startParameter.maxWorkerCount,
        )
        project.tasks.register("cleanPitestReport", Delete::class.java) { task ->
            task.delete(reportDirectory)
        }

        project.extensions.configure<PitestPluginExtension> {
            pitestVersion.set(versions.findVersion("pitest").get().requiredVersion)
            junit5PluginVersion.set(versions.findVersion("pitest-junit5-plugin").get().requiredVersion)
            testSourceSets.set(pitestTestSourceSets)
            mutators.set(setOf("STRONGER"))
            val selections = listOf(
                "mutationClasses" to targetClasses,
                "mutationTests" to targetTests,
                "mutationMutators" to mutators,
            )
            for ((name, destination) in selections) {
                val selected = project.providers.gradleProperty(name)
                if (selected.isPresent) {
                    destination.set(selected.map { value ->
                        val entries = value.split(',').map(String::trim)
                        require(entries.all { it.isNotEmpty() }) {
                            "$name must contain nonempty comma-separated values"
                        }
                        entries.toSet()
                    })
                }
            }
            threads.set(mutationWorkers)
            timeoutConstInMillis.set(10_000)
            timeoutFactor.set(BigDecimal("3.0"))
            outputFormats.set(setOf("XML", "HTML"))
            reportDir.set(reportDirectory)
            timestampedReports.set(false)
            failWhenNoMutations.set(true)
            val scratchArgument = "-Djava.io.tmpdir=" + scratchDirectory.get().asFile.absolutePath
            jvmArgs.set(listOf("-Xmx512m", "--enable-native-access=ALL-UNNAMED", scratchArgument))
            mainProcessJvmArgs.set(listOf("-Xmx512m", scratchArgument))
            jvmPath.set(
                toolchains.launcherFor { spec ->
                    spec.languageVersion.set(javaVersion.map(JavaLanguageVersion::of))
                }.map { launcher -> launcher.executablePath },
            )
        }

        val pitestTask = project.tasks.named(PitestPlugin.PITEST_TASK_NAME)
        pitestTask.configure { task ->
            // Do not delete declared outputs in a prerequisite: that defeats Gradle reuse.
            task.doFirst {
                for (directory in listOf(reportDirectory.get().asFile, scratchDirectory.get().asFile)) {
                    if (directory.exists() && !directory.deleteRecursively()) {
                        throw GradleException("Cannot clear mutation output: " + directory)
                    }
                }
                if (!scratchDirectory.get().asFile.mkdirs()) {
                    throw GradleException("Cannot create mutation scratch: " + scratchDirectory.get().asFile)
                }
            }
            task.doLast {
                if (!scratchDirectory.get().asFile.deleteRecursively()) {
                    throw GradleException("Cannot clear successful mutation scratch: " + scratchDirectory.get().asFile)
                }
            }
        }
    }
}
