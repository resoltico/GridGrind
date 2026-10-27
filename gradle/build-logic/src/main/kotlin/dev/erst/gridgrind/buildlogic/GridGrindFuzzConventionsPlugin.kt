package dev.erst.gridgrind.buildlogic

import com.diffplug.gradle.spotless.SpotlessExtension
import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.tasks.JavaExec
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
            val replay = tasks.register<Test>("fuzzReplay") {
                group = "verification"
                description = "Runs committed seeds and generator/invariant regressions without the native fuzz driver."
                testClassesDirs = fuzz.output.classesDirs
                classpath = fuzz.runtimeClasspath
                useJUnitPlatform()
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
            val harness = providers.gradleProperty("harness").orElse("").get()
            val duration = providers.gradleProperty("fuzzSeconds").orElse("60").get()
            val corpus = layout.buildDirectory.dir("fuzz/corpus/" + harness).get().asFile
            val findings = layout.buildDirectory.dir("fuzz/findings/" + harness).get().asFile
            val seeds = layout.projectDirectory.dir("src/fuzz/resources/seeds/" + harness).asFile
            tasks.register<JavaExec>("fuzz") {
                group = "verification"
                description = "Actively fuzzes one explicitly selected harness; opt-in, never part of check."
                dependsOn(agent, fuzz.classesTaskName)
                classpath = fuzz.runtimeClasspath
                mainClass.set("com.code_intelligence.jazzer.Jazzer")
                systemProperty("gridgrind.fuzz.harness", harness)
                jvmArgs("--enable-native-access=ALL-UNNAMED", "-javaagent:" + agent.flatMap { it.archiveFile }.get().asFile.absolutePath)
                args("--target_class=dev.erst.gridgrind.fuzz.FuzzTarget", "--instrumentation_includes=dev.erst.gridgrind.core.**:dev.erst.gridgrind.json.**:dev.erst.gridgrind.engine.**", "-max_total_time=" + duration, "-artifact_prefix=" + findings.absolutePath + "/", corpus.absolutePath)
                doFirst {
                    require(harness.matches(Regex("[a-z]+(?:-[a-z]+)*")) && seeds.isDirectory) { "Select -Pharness=protocol-request|protocol-workflow|engine-command-sequence|xlsx-roundtrip" }
                    require(duration.toIntOrNull()?.let { it > 0 } == true) { "fuzzSeconds must be a positive integer" }
                    corpus.mkdirs()
                    findings.mkdirs()
                    seeds.listFiles()?.filter { it.isFile }?.forEach { seed ->
                        val target = corpus.resolve(seed.name)
                        if (!target.exists()) seed.copyTo(target)
                    }
                }
            }
        }
    }
}
