package dev.erst.gridgrind.buildlogic

import java.nio.file.Path
import java.nio.file.Files
import org.gradle.testkit.runner.GradleRunner
import org.gradle.testfixtures.ProjectBuilder
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertFalse
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.Assertions.assertTrue
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.io.TempDir

/** Checks the producer's refusal boundary independently of upstream source compilation. */
class NativeProviderProducerTest {
    @TempDir lateinit var directory: Path

    @Test fun exactPatchRejectsDriftTruncationAndPathEscape() {
        val patch = "--- A.java\n+++ A.java\n@@ -1,2 +1,2 @@\n a\n-b\n+c\n"
        assertEquals("a\nc\n", NativeRepairFiles.applyPatch("a\nb\n", patch))
        assertThrows(IllegalArgumentException::class.java) { NativeRepairFiles.applyPatch("a\nx\n", patch) }
        assertThrows(IllegalArgumentException::class.java) { NativeRepairFiles.applyPatch("a\nb\n", patch.substringBefore("+c")) }
        assertThrows(IllegalArgumentException::class.java) { NativeRepairFiles.child(directory.toFile(), "../outside") }
        val file = directory.resolve("pinned").toFile().apply { writeText("actual") }
        assertThrows(IllegalArgumentException::class.java) { NativeRepairFiles.check(file, "0".repeat(64)) }
    }

    @Test fun resourceCapsuleResolvesForCompilerAndPmdRuntimeUsage() {
        Files.createDirectories(directory.resolve("gradle"))
        val repository = Path.of(System.getProperty("gridgrind.repository.root"))
        Files.copy(repository.resolve("gradle/libs.versions.toml"), directory.resolve("gradle/libs.versions.toml"))
        Files.writeString(directory.resolve("settings.gradle"), "rootProject.name = 'capsule-fixture'\ninclude 'native-provider'\n")
        Files.writeString(directory.resolve("gradle.properties"), "gridgrindJavaVersion=27\ngridgrindJavaTargetVersion=26\n")
        Files.createDirectories(directory.resolve("native-provider"))
        Files.writeString(directory.resolve("native-provider/build.gradle"), """
            plugins { id 'gridgrind.native-provider' }
            // Negative control restores the actual defect, without changing consumers.
            if (providers.gradleProperty('reintroduceUsage').isPresent()) {
                configurations.nativeCompileElements.attributes.attribute(
                    org.gradle.api.attributes.Usage.USAGE_ATTRIBUTE,
                    objects.named(org.gradle.api.attributes.Usage, org.gradle.api.attributes.Usage.JAVA_API))
            }
        """.trimIndent())
        Files.writeString(directory.resolve("build.gradle"), """
            def usages = ['compilerConsumer': 'java-api', 'pmdAuxConsumer': 'java-runtime']
            usages.each { name, requestedUsage ->
                configurations.create(name) {
                    canBeConsumed = false
                    canBeResolved = true
                    attributes.attribute(org.gradle.api.attributes.Usage.USAGE_ATTRIBUTE,
                        objects.named(org.gradle.api.attributes.Usage, requestedUsage))
                    attributes.attribute(org.gradle.api.attributes.LibraryElements.LIBRARY_ELEMENTS_ATTRIBUTE,
                        objects.named(org.gradle.api.attributes.LibraryElements, 'jar'))
                }
                dependencies.add(name, dependencies.project(path: ':native-provider', configuration: 'nativeCompileElements'))
            }
            tasks.register('resolveCapsule') {
                doLast {
                    usages.keySet().each { name ->
                        def artifacts = configurations.getByName(name).files
                        assert artifacts.size() == 1
                        assert artifacts.first().name.endsWith('.jar')
                        assert !artifacts.first().path.contains('/classes/')
                        println('CAPSULE_RESOLVED ' + name)
                    }
                }
            }
        """.trimIndent())
        fun runner(vararg extra: String) = GradleRunner.create().withProjectDir(directory.toFile()).withPluginClasspath()
            .withArguments(listOf("resolveCapsule", "--max-workers=1") + extra)
        val actual = runner().build()
        assertTrue(actual.output.contains("CAPSULE_RESOLVED compilerConsumer"), actual.output)
        assertTrue(actual.output.contains("CAPSULE_RESOLVED pmdAuxConsumer"), actual.output)
        val rejected = runner("-PreintroduceUsage=true").buildAndFail()
        assertTrue(rejected.output.contains("nativeCompileElements"), rejected.output)
        assertTrue(rejected.output.contains("runtime"), rejected.output)
    }

    @Test fun resourceListUsesActualNestedClassesAndDropsStaleOutput() {
        val project = ProjectBuilder.builder().withProjectDir(directory.toFile()).build()
        val prepared = directory.resolve("prepared").toFile().apply { mkdirs() }
        NativeRepairFiles.json(prepared.resolve("source-provenance.json"), mapOf("sources" to listOf(mapOf("sourcePath" to "org/apache/poi/Owned.java"))))
        for (role in listOf("poi", "poi-ooxml", "helper")) prepared.resolve(role).mkdirs()
        val recipes = directory.resolve("recipes").toFile().apply { mkdirs(); resolve("patches").mkdirs(); resolve("repairs.json").writeText("{}") }
        val license = directory.resolve("LICENSE").toFile().apply { writeText("license") }
        val classes = directory.resolve("classes").toFile().apply { mkdirs() }
        fun compiled(name: String) = classes.resolve("org/apache/poi/$name.class").apply { parentFile.mkdirs(); writeBytes(byteArrayOf(-54, -2, -70, -66, 0, 0, 0, 70, 0)) }
        compiled("Owned")
        val nested = compiled("Owned\$Nested")
        val task = project.tasks.create("resources", AssembleNativeProviderResources::class.java)
        task.preparedSources.set(prepared); task.recipeDirectory.set(recipes); task.recipeFiles.from(recipes)
        task.classDirectories.from(classes); task.sourceArchives.from(emptyList<Any>())
        task.apacheLicense.set(license); task.projectLicense.set(license)
        task.componentVersion.set("test"); task.compilerIdentity.set("test"); task.targetRelease.set(26)
        val output = directory.resolve("output").toFile(); task.outputDirectory.set(output)
        task.assemble()
        val list = output.resolve("META-INF/gridgrind/native-provider/required-classes.list")
        assertEquals("org.apache.poi.Owned\$Nested\norg.apache.poi.Owned\n", list.readText())
        nested.delete(); task.assemble()
        assertEquals("org.apache.poi.Owned\n", list.readText())
        assertFalse(output.resolve("META-INF/gridgrind/native-provider/classes/org/apache/poi/Owned\$Nested.class").exists())
        compiled("Unselected")
        classes.resolve("org/apache/poi/Owned.class").delete()
        assertTrue(assertThrows(IllegalArgumentException::class.java) { task.assemble() }.message!!.contains("Selected repair compiler output missing"))
    }
}
