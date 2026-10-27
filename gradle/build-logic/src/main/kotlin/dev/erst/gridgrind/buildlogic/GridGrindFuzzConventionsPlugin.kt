package dev.erst.gridgrind.buildlogic

import com.diffplug.gradle.spotless.SpotlessExtension
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.tasks.SourceSetContainer
import org.gradle.api.tasks.bundling.Jar
import org.gradle.api.tasks.testing.Test
import org.gradle.jvm.toolchain.JavaLanguageVersion
import org.gradle.kotlin.dsl.configure
import org.gradle.kotlin.dsl.getByType
import org.gradle.kotlin.dsl.register
import org.gradle.testing.jacoco.plugins.JacocoTaskExtension

/** Root-owned deterministic replay and explicitly requested standalone fuzzing. */
class GridGrindFuzzConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) {
        with(project) {
            pluginManager.apply("java")
            val libs = versionCatalog()
            extensions.getByType<JavaPluginExtension>().toolchain.languageVersion.set(
                JavaLanguageVersion.of(providers.gradleProperty("gridgrindJavaVersion").get().toInt()),
            )
            val fuzz = extensions.getByType<SourceSetContainer>().create("fuzz")
            dependencies.apply {
                add(fuzz.implementationConfigurationName, project(":engine"))
                add(fuzz.implementationConfigurationName, project(":json"))
                add(fuzz.implementationConfigurationName, platform(libs.findLibrary("junit-bom").get()))
                for (library in listOf("junit-jupiter", "jspecify", "poi-ooxml", "poi-ooxml-full", "jackson-databind", "xmlsec", "bcpkix", "bcprov", "bcutil", "slf4j-api")) {
                    add(fuzz.implementationConfigurationName, libs.findLibrary(library).get())
                }
                for (library in listOf("junit-platform-launcher", "log4j-core", "log4j-slf4j2-impl", "jazzer")) {
                    add(fuzz.runtimeOnlyConfigurationName, libs.findLibrary(library).get())
                }
            }
            configure<SpotlessExtension> {
                java {
                    it.target("src/fuzz/java/**/*.java")
                    it.googleJavaFormat(libs.findVersion("google-java-format").get().requiredVersion)
                    it.removeUnusedImports()
                    it.formatAnnotations()
                }
            }
            val replayTemporary = layout.buildDirectory.dir("tmp/fuzzReplay").get().asFile
            val replay = tasks.register<Test>("fuzzReplay") {
                group = "verification"
                description = "Runs committed seeds and generator/invariant regressions without the native fuzz driver."
                testClassesDirs = fuzz.output.classesDirs
                classpath = fuzz.runtimeClasspath
                useJUnitPlatform()
                systemProperty("java.io.tmpdir", replayTemporary.absolutePath)
                doFirst {
                    check(!replayTemporary.exists() || replayTemporary.deleteRecursively()) { "Cannot reset fuzz replay scratch: $replayTemporary" }
                    check(replayTemporary.mkdirs()) { "Cannot create fuzz replay scratch: $replayTemporary" }
                }
                systemProperty("gridgrind.fuzz.seeds", layout.projectDirectory.dir("src/fuzz/resources/seeds").asFile.absolutePath)
                inputs.dir(layout.projectDirectory.dir("src/fuzz/resources/seeds"))
                extensions.configure<JacocoTaskExtension> { isEnabled = false }
            }
            tasks.named("check") { it.dependsOn(replay) }

            // Use the pinned driver's own installer directly, without a custom premain Java class.
            val agent = tasks.register<Jar>("fuzzAgent") {
                archiveFileName.set("fuzz-agent.jar")
                destinationDirectory.set(layout.buildDirectory.dir("fuzz"))
                manifest.attributes(
                    mapOf(
                        "Premain-Class" to "com.code_intelligence.jazzer.third_party.net.bytebuddy.agent.Installer",
                        "Can-Redefine-Classes" to "true",
                        "Can-Retransform-Classes" to "true",
                        "Can-Set-Native-Method-Prefix" to "true",
                    ),
                )
            }
            val selectedHarness = providers.gradleProperty("harness").orElse("").get()
            tasks.register<FuzzTask>("fuzz") {
                group = "verification"
                description = "Actively fuzzes one explicitly selected harness; bounded, opt-in, never part of check."
                dependsOn(agent, fuzz.classesTaskName)
                classpath = fuzz.runtimeClasspath
                mainClass.set("com.code_intelligence.jazzer.Jazzer")
                jvmArgs("--enable-native-access=ALL-UNNAMED", "-javaagent:" + agent.flatMap { it.archiveFile }.get().asFile.absolutePath)
                harness.set(selectedHarness)
                seconds.set(providers.gradleProperty("fuzzSeconds").orElse("60").map(String::toInt))
                runs.set(providers.gradleProperty("fuzzRuns").orElse("10000").map(String::toInt))
                maxInputBytes.set(providers.gradleProperty("fuzzMaxInputBytes").orElse("32768").map(String::toInt))
                seeds.set(layout.projectDirectory.dir("src/fuzz/resources/seeds/" + selectedHarness))
                workspace.set(layout.buildDirectory.dir("fuzz"))
            }
        }
    }
}
