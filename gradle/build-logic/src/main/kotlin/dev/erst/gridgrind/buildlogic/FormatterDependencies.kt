package dev.erst.gridgrind.buildlogic

import org.gradle.api.Project

/** Aligns public libraries inside Spotless's detached formatter provisioning graph. */
internal fun Project.configureFormatterDependencies() {
    val catalog = versionCatalog()
    val formatterVersion = catalog.findVersion("google-java-format").get().requiredVersion
    val selected = listOf("guava", "errorprone-annotations", "jspecify").associate { name ->
        val dependency = catalog.findLibrary(name).get().get()
        "${dependency.module.group}:${dependency.module.name}" to dependency.versionConstraint.requiredVersion
    }
    dependencies.components.withModule("com.google.googlejavaformat:google-java-format") {
        // The compiler's separately owned formatter API remains unchanged.
        if (id.version == formatterVersion) {
            allVariants {
                withDependencies {
                    forEach { dependency ->
                        selected["${dependency.group}:${dependency.name}"]?.let { version ->
                            dependency.version { require(version) }
                        }
                    }
                }
            }
        }
    }
}
