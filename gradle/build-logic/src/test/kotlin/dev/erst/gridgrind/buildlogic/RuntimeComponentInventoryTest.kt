package dev.erst.gridgrind.buildlogic

import groovy.json.JsonSlurper
import java.nio.file.Path
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream
import org.gradle.api.GradleException
import org.gradle.testfixtures.ProjectBuilder
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertFalse
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.io.TempDir

/** Checks resolver identity, real capsule ownership and separation of retained/local inventories. */
class RuntimeComponentInventoryTest {
    @TempDir lateinit var root: Path

    @Test fun usesResolvedCoordinatesRatherThanMisleadingFilenames() {
        val external = root.resolve("misleading-99.jar").toFile().apply { writeText("actual runtime bytes") }
        val task = task("inventory", "0.78.0", "0.78.0")
        task.components.set(listOf(RuntimeMavenComponent("example", "actual", "1.2", external)))
        task.writeInventory()
        val retained = JsonSlurper().parse(task.inventoryFile.get().asFile) as Map<*, *>
        val rows = retained["components"] as List<*>
        assertEquals(2, rows.size)
        assertEquals(listOf("LICENSE", "NOTICE"), retained["legalFiles"])
        assertEquals("actual", (rows[1] as Map<*, *>)["artifact"])
        assertEquals("1.2", (rows[1] as Map<*, *>)["version"])
        assertFalse(task.inventoryFile.get().asFile.readText().contains("misleading"))
        assertFalse(task.inventoryFile.get().asFile.readText().contains(root.toString()))
        val local = JsonSlurper().parse(task.resolvedFilesFile.get().asFile) as Map<*, *>
        assertEquals(external.absolutePath, ((local["components"] as List<*>)[1] as Map<*, *>)["file"])
    }

    @Test fun refusesAStaleOrInventedProjectIdentityAndEmptyExternalGraph() {
        val external = root.resolve("library.jar").toFile().apply { writeText("runtime") }
        val stale = task("stale", "0.78.0", "0.77.0")
        stale.components.set(listOf(RuntimeMavenComponent("example", "library", "1.0", external)))
        assertThrows(GradleException::class.java) { stale.writeInventory() }
        val wrongMaven = task("wrongMaven", "0.78.0", "0.78.0", "0.77.0")
        wrongMaven.components.set(listOf(RuntimeMavenComponent("example", "library", "1.0", external)))
        assertThrows(GradleException::class.java) { wrongMaven.writeInventory() }
        val empty = task("empty", "0.78.0", "0.78.0")
        empty.components.set(emptyList())
        assertThrows(GradleException::class.java) { empty.writeInventory() }
    }

    private fun task(
        name: String, expectedVersion: String, packagedVersion: String, mavenVersion: String = packagedVersion,
    ): WriteRuntimeComponentInventoryTask {
        val capsule = root.resolve("$name-capsule.jar").toFile()
        ZipOutputStream(capsule.outputStream()).use { zip ->
            for ((path, value) in mapOf(
                "META-INF/gridgrind/native-provider/manifest.json" to """{"component":{"group":"dev.erst.gridgrind","artifact":"native-provider","version":"$packagedVersion"}}""",
                "META-INF/maven/dev.erst.gridgrind/native-provider/pom.properties" to "groupId=dev.erst.gridgrind\nartifactId=native-provider\nversion=$mavenVersion\n",
            )) {
                zip.putNextEntry(ZipEntry(path)); zip.write(value.toByteArray(Charsets.UTF_8)); zip.closeEntry()
            }
        }
        val project = ProjectBuilder.builder().withProjectDir(root.toFile()).build()
        return project.tasks.register(name, WriteRuntimeComponentInventoryTask::class.java).get().apply {
            legalFiles.set(listOf("LICENSE", "NOTICE"))
            nativeGroup.set("dev.erst.gridgrind"); nativeArtifact.set("native-provider")
            nativeVersion.set(expectedVersion); nativeCapsule.set(capsule)
            inventoryFile.set(root.resolve("$name-inventory.json").toFile())
            resolvedFilesFile.set(root.resolve("$name-files.json").toFile())
        }
    }
}
