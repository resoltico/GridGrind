import java.nio.file.DirectoryNotEmptyException
import java.nio.file.Files
import java.util.Comparator
import org.gradle.api.DefaultTask
import org.gradle.api.file.DirectoryProperty
import org.gradle.api.plugins.JavaPluginExtension
import org.gradle.api.tasks.SourceSetContainer
import org.gradle.api.tasks.InputDirectory
import org.gradle.api.tasks.Optional
import org.gradle.api.tasks.TaskAction
import org.gradle.api.tasks.JavaExec
import org.gradle.api.tasks.testing.Test
import org.gradle.kotlin.dsl.get
import org.gradle.kotlin.dsl.named
import org.gradle.kotlin.dsl.register

plugins {
    java
}

description = "Local-only Jazzer fuzzing layer for GridGrind"

repositories {
    mavenCentral()
}

abstract class CleanLocalCorpusTask : DefaultTask() {
    @get:InputDirectory
    @get:Optional
    abstract val localDirectory: DirectoryProperty

    @TaskAction
    fun clean() {
        val localPath = localDirectory.asFile.orNull?.toPath() ?: return
        if (!Files.exists(localPath)) {
            return
        }
        Files.walk(localPath).use { localStream ->
            localStream
                .filter { path -> path.fileName.toString() == ".cifuzz-corpus" }
                .sorted(Comparator.reverseOrder())
                .forEach { corpusPath ->
                    Files.walk(corpusPath).use { corpusStream ->
                        corpusStream.sorted(Comparator.reverseOrder()).forEach(Files::deleteIfExists)
                    }
                }
        }
    }
}

abstract class CleanLocalFindingsTask : DefaultTask() {
    @get:InputDirectory
    @get:Optional
    abstract val runsDirectory: DirectoryProperty

    @TaskAction
    fun clean() {
        val runsPath = runsDirectory.asFile.orNull?.toPath() ?: return
        if (!Files.exists(runsPath)) {
            return
        }
        Files.walk(runsPath).use { runsStream ->
            runsStream.sorted(Comparator.reverseOrder()).forEach { path ->
                if (path == runsPath) {
                    return@forEach
                }
                if (path.fileName.toString() == ".cifuzz-corpus") {
                    return@forEach
                }
                if (path.iterator().asSequence().any { it.toString() == ".cifuzz-corpus" }) {
                    return@forEach
                }
                try {
                    Files.deleteIfExists(path)
                } catch (_: DirectoryNotEmptyException) {
                    // Preserved corpus content intentionally keeps some run directories alive.
                }
            }
        }
    }
}

val gridgrindJavaVersion: Int =
    providers.gradleProperty("gridgrindJavaVersion").map(String::toInt).get()
val jazzerMaxDuration = providers.gradleProperty("jazzerMaxDuration").orNull
val jazzerMaxExecutions = providers.gradleProperty("jazzerMaxExecutions").orNull

extensions.configure<JavaPluginExtension> {
    toolchain {
        languageVersion = JavaLanguageVersion.of(gridgrindJavaVersion)
    }
}

val sourceSets = extensions.getByType(SourceSetContainer::class)
val mainSourceSet = sourceSets["main"]
val fuzzSourceSet = sourceSets.create("fuzz") {
    java.setSrcDirs(listOf("src/fuzz/java"))
    resources.setSrcDirs(listOf("src/fuzz/resources"))
}
fuzzSourceSet.compileClasspath += mainSourceSet.output
fuzzSourceSet.runtimeClasspath += mainSourceSet.output

configurations.named(fuzzSourceSet.implementationConfigurationName) {
    extendsFrom(configurations["implementation"])
}
configurations.named(fuzzSourceSet.runtimeOnlyConfigurationName) {
    extendsFrom(configurations["runtimeOnly"])
}

dependencies {
    add("implementation", platform("org.junit:junit-bom:6.0.3"))
    add("implementation", "org.junit.platform:junit-platform-launcher")
    add("implementation", "com.code-intelligence:jazzer-api:0.30.0")
    add("implementation", "com.code-intelligence:jazzer:0.30.0")
    add("implementation", "org.apache.poi:poi-ooxml:5.5.1")
    add("implementation", "tools.jackson.core:jackson-databind:3.0.3")
    add("implementation", "dev.erst.gridgrind:protocol")
    add("implementation", "dev.erst.gridgrind:engine")
    add("runtimeOnly", "org.apache.logging.log4j:log4j-core:2.25.3")
    add("testImplementation", platform("org.junit:junit-bom:6.0.3"))
    add("testImplementation", "org.junit.jupiter:junit-jupiter")
    add("testImplementation", "org.apache.poi:poi-ooxml:5.5.1")
    add("testImplementation", "tools.jackson.core:jackson-databind:3.0.3")
    add("testImplementation", "dev.erst.gridgrind:protocol")
    add("testImplementation", "dev.erst.gridgrind:engine")
    add("testRuntimeOnly", "org.junit.platform:junit-platform-launcher")
    add(fuzzSourceSet.implementationConfigurationName, platform("org.junit:junit-bom:6.0.3"))
    add(fuzzSourceSet.implementationConfigurationName, "org.junit.jupiter:junit-jupiter")
    add(fuzzSourceSet.runtimeOnlyConfigurationName, "org.junit.platform:junit-platform-launcher")
    add(fuzzSourceSet.implementationConfigurationName, "com.code-intelligence:jazzer-junit:0.30.0")
    add(fuzzSourceSet.implementationConfigurationName, "com.code-intelligence:jazzer-api:0.30.0")
    add(fuzzSourceSet.implementationConfigurationName, "org.apache.poi:poi-ooxml:5.5.1")
    add(fuzzSourceSet.implementationConfigurationName, "dev.erst.gridgrind:protocol")
    add(fuzzSourceSet.implementationConfigurationName, "dev.erst.gridgrind:engine")
}

fun runDirectory(path: String) = layout.projectDirectory.dir(path).asFile

fun requiredGradleProperty(name: String): String =
    providers.gradleProperty(name).orNull
        ?: throw IllegalArgumentException("Missing Gradle property: $name")

fun JavaExec.configureHarnessRuntime() {
    classpath = fuzzSourceSet.runtimeClasspath
    mainClass.set("dev.erst.gridgrind.jazzer.tool.JazzerHarnessRunner")
    outputs.upToDateWhen { false }
    workingDir = layout.projectDirectory.asFile
    jvmArgs("--enable-native-access=ALL-UNNAMED")
    if (jazzerMaxDuration != null) {
        systemProperty("jazzer.max_duration", jazzerMaxDuration)
    }
    if (jazzerMaxExecutions != null) {
        systemProperty("jazzer.max_executions", jazzerMaxExecutions)
    }
}

fun JavaExec.configureMainSourceSet() {
    classpath = sourceSets["main"].runtimeClasspath
    mainClass.set("dev.erst.gridgrind.jazzer.tool.JazzerCli")
    workingDir = layout.projectDirectory.asFile
    jvmArgs("--enable-native-access=ALL-UNNAMED")
}

fun registerToolTask(
    name: String,
    descriptionText: String,
    argumentsProvider: () -> List<String>
) = tasks.register<JavaExec>(name) {
    description = descriptionText
    group = "verification"
    configureMainSourceSet()
    notCompatibleWithConfigurationCache(
        "Local-only Jazzer tool tasks assemble command arguments at execution time."
    )
    doFirst {
        setArgs(argumentsProvider())
    }
}

fun registerFuzzTask(
    name: String,
    descriptionText: String,
    className: String,
    workingDirectory: String,
    fuzzing: Boolean
) = tasks.register<JavaExec>(name) {
    description = descriptionText
    group = "verification"
    configureHarnessRuntime()
    args("--class", className)
    workingDir = runDirectory(workingDirectory)
    doFirst {
        workingDir.mkdirs()
    }
    if (fuzzing) {
        environment("JAZZER_FUZZ", "1")
    }
}

fun registerRegressionTask(
    name: String,
    descriptionText: String,
    className: String,
    workingDirectory: String
) = tasks.register<JavaExec>(name) {
    description = descriptionText
    group = "verification"
    configureHarnessRuntime()
    args("--class", className)
    workingDir = runDirectory(workingDirectory)
    doFirst {
        workingDir.mkdirs()
    }
}

val regressionProtocolRequest =
    registerRegressionTask(
        "regressionProtocolRequest",
        "Replays committed protocol-request Jazzer inputs in regression mode.",
        "dev.erst.gridgrind.jazzer.protocol.ProtocolRequestFuzzTest",
        ".local/runs/regression/protocol-request"
    )

val regressionProtocolWorkflow =
    registerRegressionTask(
        "regressionProtocolWorkflow",
        "Replays committed protocol-workflow Jazzer inputs in regression mode.",
        "dev.erst.gridgrind.jazzer.protocol.OperationWorkflowFuzzTest",
        ".local/runs/regression/protocol-workflow"
    )

val regressionEngineCommandSequence =
    registerRegressionTask(
        "regressionEngineCommandSequence",
        "Replays committed engine-command-sequence Jazzer inputs in regression mode.",
        "dev.erst.gridgrind.jazzer.engine.WorkbookCommandSequenceFuzzTest",
        ".local/runs/regression/engine-command-sequence"
    )

val regressionXlsxRoundTrip =
    registerRegressionTask(
        "regressionXlsxRoundTrip",
        "Replays committed xlsx-roundtrip Jazzer inputs in regression mode.",
        "dev.erst.gridgrind.jazzer.engine.XlsxRoundTripFuzzTest",
        ".local/runs/regression/xlsx-roundtrip"
    )

val jazzerRegression =
    tasks.register("jazzerRegression") {
        description = "Runs all Jazzer harnesses in regression mode."
        group = "verification"
        dependsOn(
            regressionProtocolRequest,
            regressionProtocolWorkflow,
            regressionEngineCommandSequence,
            regressionXlsxRoundTrip
        )
    }

val fuzzProtocolRequest =
    registerFuzzTask(
        "fuzzProtocolRequest",
        "Actively fuzzes protocol request parsing and validation.",
        "dev.erst.gridgrind.jazzer.protocol.ProtocolRequestFuzzTest",
        ".local/runs/protocol-request",
        true
    )

val fuzzProtocolWorkflow =
    registerFuzzTask(
        "fuzzProtocolWorkflow",
        "Actively fuzzes ordered GridGrind protocol workflows.",
        "dev.erst.gridgrind.jazzer.protocol.OperationWorkflowFuzzTest",
        ".local/runs/protocol-workflow",
        true
    )

val fuzzEngineCommandSequence =
    registerFuzzTask(
        "fuzzEngineCommandSequence",
        "Actively fuzzes workbook-core command sequences.",
        "dev.erst.gridgrind.jazzer.engine.WorkbookCommandSequenceFuzzTest",
        ".local/runs/engine-command-sequence",
        true
    )

val fuzzXlsxRoundTrip =
    registerFuzzTask(
        "fuzzXlsxRoundTrip",
        "Actively fuzzes .xlsx save and reopen invariants.",
        "dev.erst.gridgrind.jazzer.engine.XlsxRoundTripFuzzTest",
        ".local/runs/xlsx-roundtrip",
        true
    )

tasks.named<Test>("test") {
    description = "Runs deterministic Jazzer support tests."
    group = "verification"
    useJUnitPlatform()
    maxParallelForks = 1
    jvmArgs("--enable-native-access=ALL-UNNAMED")
}

tasks.named("check") {
    dependsOn(jazzerRegression)
}

tasks.register("fuzzAllLocal") {
    description = "Runs all active local-only fuzzing tasks."
    group = "verification"
    dependsOn(
        fuzzProtocolRequest,
        fuzzProtocolWorkflow,
        fuzzEngineCommandSequence,
        fuzzXlsxRoundTrip
    )
}

fuzzProtocolWorkflow.configure {
    mustRunAfter(fuzzProtocolRequest)
}

fuzzEngineCommandSequence.configure {
    mustRunAfter(fuzzProtocolWorkflow)
}

fuzzXlsxRoundTrip.configure {
    mustRunAfter(fuzzEngineCommandSequence)
}

val jazzerSupportTests = tasks.named<Test>("test")

regressionProtocolRequest.configure {
    mustRunAfter(jazzerSupportTests)
}

regressionProtocolWorkflow.configure {
    mustRunAfter(regressionProtocolRequest)
    mustRunAfter(jazzerSupportTests)
}

regressionEngineCommandSequence.configure {
    mustRunAfter(regressionProtocolWorkflow)
    mustRunAfter(jazzerSupportTests)
}

regressionXlsxRoundTrip.configure {
    mustRunAfter(regressionEngineCommandSequence)
    mustRunAfter(jazzerSupportTests)
}

tasks.register<CleanLocalCorpusTask>("cleanLocalCorpus") {
    description = "Deletes generated Jazzer corpora under .local."
    group = "build"
    localDirectory.set(layout.projectDirectory.dir(".local"))
}

tasks.register<CleanLocalFindingsTask>("cleanLocalFindings") {
    description = "Deletes local crash files and non-corpus run state under .local."
    group = "build"
    runsDirectory.set(layout.projectDirectory.dir(".local/runs"))
}

registerToolTask(
    "jazzerSummarizeRun",
    "Builds the latest summary artifacts for one completed Jazzer run."
) {
    listOf(
        "summarize-run",
        "--project-dir",
        layout.projectDirectory.asFile.absolutePath,
        "--target",
        requiredGradleProperty("jazzerTarget"),
        "--task",
        requiredGradleProperty("jazzerTaskName"),
        "--log",
        requiredGradleProperty("jazzerLogPath"),
        "--history",
        requiredGradleProperty("jazzerHistoryDirectory"),
        "--started-at",
        requiredGradleProperty("jazzerStartedAt"),
        "--finished-at",
        requiredGradleProperty("jazzerFinishedAt"),
        "--exit-code",
        requiredGradleProperty("jazzerExitCode"),
        "--corpus-before-files",
        requiredGradleProperty("jazzerCorpusBeforeFiles"),
        "--corpus-before-bytes",
        requiredGradleProperty("jazzerCorpusBeforeBytes")
    )
}

registerToolTask("jazzerStatus", "Shows a concise latest-summary view across Jazzer targets.") {
    buildList {
        add("status")
        add("--project-dir")
        add(layout.projectDirectory.asFile.absolutePath)
        providers.gradleProperty("jazzerTarget").orNull?.let {
            add("--target")
            add(it)
        }
    }
}

registerToolTask("jazzerReport", "Shows the detailed latest-summary report for Jazzer targets.") {
    buildList {
        add("report")
        add("--project-dir")
        add(layout.projectDirectory.asFile.absolutePath)
        providers.gradleProperty("jazzerTarget").orNull?.let {
            add("--target")
            add(it)
        }
    }
}

registerToolTask("jazzerListFindings", "Lists replayed local finding artifacts.") {
    buildList {
        add("list-findings")
        add("--project-dir")
        add(layout.projectDirectory.asFile.absolutePath)
        providers.gradleProperty("jazzerTarget").orNull?.let {
            add("--target")
            add(it)
        }
    }
}

registerToolTask("jazzerListCorpus", "Lists local corpus state and promoted inputs.") {
    buildList {
        add("list-corpus")
        add("--project-dir")
        add(layout.projectDirectory.asFile.absolutePath)
        providers.gradleProperty("jazzerTarget").orNull?.let {
            add("--target")
            add(it)
        }
    }
}

registerToolTask("jazzerReplay", "Replays one local input against a single Jazzer harness.") {
    buildList {
        add("replay")
        add("--project-dir")
        add(layout.projectDirectory.asFile.absolutePath)
        add("--target")
        add(requiredGradleProperty("jazzerTarget"))
        add("--input")
        add(requiredGradleProperty("jazzerInput"))
        if (providers.gradleProperty("jazzerJsonOutput").orNull == "true") {
            add("--json")
        }
    }
}

registerToolTask("jazzerPromote", "Promotes one local input into committed Jazzer regression resources.") {
    listOf(
        "promote",
        "--project-dir",
        layout.projectDirectory.asFile.absolutePath,
        "--target",
        requiredGradleProperty("jazzerTarget"),
        "--input",
        requiredGradleProperty("jazzerInput"),
        "--name",
        requiredGradleProperty("jazzerName")
    )
}
