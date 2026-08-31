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
import org.gradle.kotlin.dsl.create
import org.gradle.kotlin.dsl.getByType

/** Shared PIT wiring for explicitly scoped product modules. */
class GridGrindMutationConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        project.pluginManager.apply(PitestPlugin::class.java)
        val scope = project.extensions.create<GridGrindMutationScopeExtension>("gridgrindMutation")
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
            targetClasses.set(scope.targetClasses)
            targetTests.set(scope.targetTests)
            testSourceSets.set(pitestTestSourceSets)
            mutationThreshold.set(scope.mutationThreshold)
            coverageThreshold.set(scope.coverageThreshold)
            testStrengthThreshold.set(scope.testStrengthThreshold)
            maxSurviving.set(scope.maxSurviving)
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

        val verifyScope =
            project.tasks.register("verifyPitestScope", VerifyPitestScopeTask::class.java) { task ->
                task.group = "verification"
                task.description =
                    "Verifies that every configured PIT class and test pattern resolves to compiled bytecode."
                task.targetClassPatterns.set(scope.targetClasses)
                task.targetTestPatterns.set(scope.targetTests)
                task.productionClassDirectories.from(
                    sourceSets.named("main").map { sourceSet -> sourceSet.output.classesDirs },
                )
                task.testClassDirectories.from(
                    pitestTestSourceSets.map { sourceSet -> sourceSet.output.classesDirs },
                )
                task.reportFile.set(reportDirectory.map { directory -> directory.file("scope.tsv") })
                task.dependsOn("classes")
                task.dependsOn(pitestTestSourceSets.map { sourceSet -> sourceSet.classesTaskName })
                task.mustRunAfter(cleanReport)
            }
        val verifyReport =
            project.tasks.register("verifyPitestReport", VerifyPitestReportTask::class.java) { task ->
                task.group = "verification"
                task.description =
                    "Rejects every PIT result other than a killed mutant."
                task.mutationReport.set(reportDirectory.map { directory -> directory.file("mutations.xml") })
                task.verificationReport.set(
                    reportDirectory.map { directory -> directory.file("verification.tsv") },
                )
            }

        val pitestTask = project.tasks.named(PitestPlugin.PITEST_TASK_NAME)
        pitestTask.configure { task ->
            task.dependsOn(cleanReport)
            task.dependsOn(verifyScope)
        }
        verifyReport.configure { task -> task.dependsOn(pitestTask) }
    }
}
