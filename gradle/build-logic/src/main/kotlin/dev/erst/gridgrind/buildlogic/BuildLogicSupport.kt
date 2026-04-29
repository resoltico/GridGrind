package dev.erst.gridgrind.buildlogic

import org.gradle.api.Action
import org.gradle.api.Project
import org.gradle.api.artifacts.VersionCatalog
import org.gradle.api.artifacts.VersionCatalogsExtension
import org.gradle.api.artifacts.repositories.MavenArtifactRepository
import org.gradle.kotlin.dsl.getByType

private const val JACOCO_SNAPSHOTS_REPOSITORY =
    "https://central.sonatype.com/repository/maven-snapshots/"

internal fun Project.versionCatalog(name: String = "libs"): VersionCatalog =
    extensions.getByType<VersionCatalogsExtension>().named(name)

internal fun Project.configureGridGrindRepositories() {
    repositories.mavenCentral()
    repositories.maven(
        Action<MavenArtifactRepository> { repository ->
            repository.name = "JaCoCoSnapshots"
            repository.url = uri(JACOCO_SNAPSHOTS_REPOSITORY)
            repository.mavenContent { content ->
                content.snapshotsOnly()
                content.includeGroup("org.jacoco")
            }
        },
    )
}
