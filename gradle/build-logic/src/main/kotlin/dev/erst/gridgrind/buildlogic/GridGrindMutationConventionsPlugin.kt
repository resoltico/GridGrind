package dev.erst.gridgrind.buildlogic

import info.solidsoft.gradle.pitest.PitestPlugin
import info.solidsoft.gradle.pitest.PitestPluginExtension
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
        val cleanReport =
            project.tasks.register("cleanPitestReport", Delete::class.java) { task ->
                task.delete(reportDirectory)
            }

        project.extensions.configure<PitestPluginExtension> {
            pitestVersion.set(versions.findVersion("pitest").get().requiredVersion)
            junit5PluginVersion.set(versions.findVersion("pitest-junit5-plugin").get().requiredVersion)
            testSourceSets.set(pitestTestSourceSets)
            mutators.set(setOf("STRONGER"))
            threads.set(4)
            timeoutConstInMillis.set(10_000)
            timeoutFactor.set(BigDecimal("3.0"))
            outputFormats.set(setOf("XML", "HTML"))
            reportDir.set(reportDirectory)
            timestampedReports.set(false)
            failWhenNoMutations.set(true)
            jvmArgs.set(listOf("--enable-native-access=ALL-UNNAMED"))
            jvmPath.set(
                toolchains.launcherFor { spec ->
                    spec.languageVersion.set(javaVersion.map(JavaLanguageVersion::of))
                }.map { launcher -> launcher.executablePath },
            )
        }

        val pitestTask = project.tasks.named(PitestPlugin.PITEST_TASK_NAME)
        pitestTask.configure { task ->
            task.dependsOn(cleanReport)
        }
    }
}
