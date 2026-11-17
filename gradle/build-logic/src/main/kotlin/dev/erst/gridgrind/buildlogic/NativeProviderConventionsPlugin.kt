package dev.erst.gridgrind.buildlogic

import org.gradle.api.Plugin
import org.gradle.api.Project
import org.gradle.api.attributes.Category
import org.gradle.api.attributes.LibraryElements
import org.gradle.api.attributes.Usage
import org.gradle.api.tasks.bundling.Jar
import org.gradle.api.tasks.compile.JavaCompile
import org.gradle.jvm.toolchain.JavaLanguageVersion
import org.gradle.jvm.toolchain.JavaToolchainService
import org.gradle.jvm.toolchain.JvmVendorSpec
import org.gradle.kotlin.dsl.getByType
import org.gradle.kotlin.dsl.register

/** Explicit producer for reviewed native compiler views and a resource-only delivery capsule. */
class NativeProviderConventionsPlugin : Plugin<Project> {
    override fun apply(project: Project) = with(project) {
        pluginManager.apply("base")
        pluginManager.apply("jvm-toolchains")
        configureGridGrindRepositories()
        val catalog = versionCatalog()
        val release = providers.gradleProperty("gridgrindJavaTargetVersion").map(String::toInt)
        val major = providers.gradleProperty("gridgrindJavaVersion").map(String::toInt)
        val compiler = extensions.getByType<JavaToolchainService>().compilerFor {
            languageVersion.set(major.map(JavaLanguageVersion::of))
            vendor.set(JvmVendorSpec.ADOPTIUM)
        }
        val vendorClasspath = configurations.create("vendorClasspath") {
            isCanBeConsumed = false; isCanBeResolved = true
        }
        val poiSources = configurations.create("poiSources") {
            isCanBeConsumed = false; isCanBeResolved = true; isTransitive = false
        }
        val ooxmlSources = configurations.create("ooxmlSources") {
            isCanBeConsumed = false; isCanBeResolved = true; isTransitive = false
        }
        val poiVersion = catalog.findVersion("poi-ooxml").get().requiredVersion
        dependencies.add(poiSources.name, "org.apache.poi:poi:$poiVersion:sources@jar")
        dependencies.add(ooxmlSources.name, "org.apache.poi:poi-ooxml:$poiVersion:sources@jar")
        for (name in listOf("poi-ooxml", "poi-ooxml-full")) dependencies.add(vendorClasspath.name, catalog.findLibrary(name).get())
        dependencies.constraints {
            for (name in listOf("bnd-annotations", "osgi-resource", "commons-codec", "commons-lang3", "woodstox-core", "stax2-api", "slf4j-api", "jspecify", "errorprone-annotations")) {
                add(vendorClasspath.name, catalog.findLibrary(name).get().get())
            }
        }
        val prepare = tasks.register<PrepareNativeProviderSources>("prepareNativeProviderSources") {
            recipeDirectory.set(layout.projectDirectory)
            recipeFiles.from(fileTree(layout.projectDirectory) { include("repairs.json", "patches/**", "src/helper/java/**") })
            compilerIdentity.set(compiler.map { "${it.metadata.vendor}:${it.metadata.javaRuntimeVersion}" })
            requiredCompilerRuntime.set(providers.fileContents(rootProject.layout.projectDirectory.file("release/container-runtime.properties")).asText.map { text ->
                java.util.Properties().apply { text.reader().use(::load) }.getProperty("TEMURIN_VERSION")
            })
            this.poiSources.set(layout.file(poiSources.elements.map { it.single().asFile }))
            this.ooxmlSources.set(layout.file(ooxmlSources.elements.map { it.single().asFile }))
            outputDirectory.set(layout.buildDirectory.dir("generated/provider-sources"))
        }
        fun compile(name: String, role: String) = tasks.register<JavaCompile>(name) {
            source(prepare.flatMap { it.outputDirectory.dir(role) })
            classpath = vendorClasspath
            javaCompiler.set(compiler)
            options.release.set(release)
            sourceCompatibility = release.get().toString()
            targetCompatibility = release.get().toString()
            options.encoding = "UTF-8"
            options.compilerArgs.add("-implicit:none")
            modularity.inferModulePath.set(false)
            destinationDirectory.set(layout.buildDirectory.dir("classes/provider/$role"))
        }
        val helper = compile("compileNativeHelper", "helper")
        helper.configure { options.compilerArgs.addAll(listOf("-Xlint:all", "-Werror")) }
        val poi = compile("compilePoiRepairs", "poi")
        poi.configure { classpath = files(helper.flatMap { it.destinationDirectory }) + vendorClasspath }
        val ooxml = compile("compileOoxmlRepairs", "poi-ooxml")
        ooxml.configure { classpath = files(helper.flatMap { it.destinationDirectory }, poi.flatMap { it.destinationDirectory }) + vendorClasspath }
        val ownVersion = version.toString()
        val resources = tasks.register<AssembleNativeProviderResources>("assembleNativeProviderResources") {
            preparedSources.set(prepare.flatMap { it.outputDirectory })
            recipeDirectory.set(layout.projectDirectory)
            recipeFiles.from(fileTree(layout.projectDirectory) { include("repairs.json", "patches/**", "src/helper/java/**") })
            classDirectories.from(helper.flatMap { it.destinationDirectory }, poi.flatMap { it.destinationDirectory }, ooxml.flatMap { it.destinationDirectory })
            sourceArchives.from(poiSources, ooxmlSources)
            apacheLicense.set(rootProject.layout.projectDirectory.file("LICENSE-APACHE-2.0"))
            projectLicense.set(rootProject.layout.projectDirectory.file("LICENSE"))
            componentVersion.set(ownVersion)
            compilerIdentity.set(compiler.map { "${it.metadata.vendor}:${it.metadata.javaRuntimeVersion}" })
            targetRelease.set(release)
            outputDirectory.set(layout.buildDirectory.dir("generated/provider-resources"))
        }
        val capsule = tasks.register<Jar>("nativeProviderJar") {
            from(resources.flatMap { it.outputDirectory })
            archiveBaseName.set("gridgrind-native-provider")
            archiveVersion.set(ownVersion)
            destinationDirectory.set(layout.buildDirectory.dir("libs"))
            isPreserveFileTimestamps = false; isReproducibleFileOrder = true
            manifest.attributes(mapOf("Automatic-Module-Name" to "dev.erst.gridgrind.nativeprovider", "Implementation-Title" to "GridGrind native provider repairs", "Implementation-Version" to ownVersion, "Implementation-License" to "Apache-2.0 AND AGPL-3.0-only"))
        }
        fun variant(name: String, usage: String?, elements: String) = configurations.create(name) {
            isCanBeConsumed = true; isCanBeResolved = false
            if (usage != null) attributes.attribute(Usage.USAGE_ATTRIBUTE, objects.named(Usage::class.java, usage))
            attributes.attribute(Category.CATEGORY_ATTRIBUTE, objects.named(Category::class.java, Category.LIBRARY))
            attributes.attribute(LibraryElements.LIBRARY_ELEMENTS_ATTRIBUTE, objects.named(LibraryElements::class.java, elements))
        }
        variant("poiCompileElements", Usage.JAVA_API, LibraryElements.CLASSES).apply {
            attributes.attribute(NativeProviderVariants.REPAIR_MODULE, NativeProviderVariants.POI_MODULE)
            outgoing.artifact(helper.flatMap { it.destinationDirectory }); outgoing.artifact(poi.flatMap { it.destinationDirectory })
        }
        variant("ooxmlCompileElements", Usage.JAVA_API, LibraryElements.CLASSES).apply {
            attributes.attribute(NativeProviderVariants.REPAIR_MODULE, NativeProviderVariants.OOXML_MODULE)
            outgoing.artifact(ooxml.flatMap { it.destinationDirectory })
        }
        // This resource-only capsule has identical compile/runtime content. PMD auxiliary
        // classpaths inherit compileOnly dependencies while requesting JAVA_RUNTIME usage.
        variant("nativeCompileElements", null, LibraryElements.JAR).outgoing.artifact(capsule)
        variant("nativeRuntimeElements", Usage.JAVA_RUNTIME, LibraryElements.JAR).outgoing.artifact(capsule)
        tasks.named("assemble") { dependsOn(capsule) }
        Unit
    }
}
