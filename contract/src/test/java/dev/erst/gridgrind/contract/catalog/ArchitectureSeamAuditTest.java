package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.regex.Pattern;
import org.junit.jupiter.api.Test;

/** Build-failing architectural audits over key module and seam boundaries. */
class ArchitectureSeamAuditTest {
  private static final Pattern QUALIFIED_ENGINE_EXPORT_PATTERN =
      Pattern.compile("exports dev\\.erst\\.gridgrind\\.engine\\.api;");
  private static final String STALE_MODULE_GRAPH = "`cli -> protocol -> engine`";
  private static final String STALE_JAZZER_MODULES = "local `engine` and `protocol` modules";
  private static final String STALE_NULL_DEFAULT_GUIDANCE = "The sole sanctioned null-return site";
  private static final String STALE_PMD_PROTOCOL_ORCHESTRATOR = "GridGrindService";
  private static final String STALE_PMD_EXCEPTION_HELPER = "withExceptionData()";
  private static final String STALE_EXECUTOR_RUNTIME_GUIDANCE =
      "executor/   The only execution bridge from the canonical contract into the";
  private static final String STALE_CLI_EXECUTOR_GUIDANCE =
      "or --request file, delegates to executor, writes the response,";
  private static final String STALE_KOTLIN_PROTOCOL_LOAD = ".codex/AGENTS_KOTLIN24_GRADLE.md";
  private static final String CURRENT_SQLITE_SURFACE =
      "SQLite surface touched: build, link, SQL, migrations, WAL, durability, bindings (baseline 3.53.2)";

  @Test
  void pmdRulesetCommentsStayCurrent() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    String ruleset = Files.readString(repositoryRoot.resolve("gradle/pmd/ruleset.xml"));

    assertTrue(
        !ruleset.contains(STALE_PMD_PROTOCOL_ORCHESTRATOR),
        "gradle/pmd/ruleset.xml must not justify exclusions with deleted GridGrindService");
    assertTrue(
        !ruleset.contains(STALE_PMD_EXCEPTION_HELPER),
        "gradle/pmd/ruleset.xml must not justify exclusions with deleted withExceptionData()");
    assertTrue(
        ruleset.contains("verifyJavaSourceShape"),
        "gradle/pmd/ruleset.xml must point at the build-owned source-shape gate");
    assertTrue(
        ruleset.contains("verifyControlPlaneShape"),
        "gradle/pmd/ruleset.xml must point at the build-owned control-plane shape gate");
    assertTrue(
        ruleset.contains("verifyJavaSourceDuplication"),
        "gradle/pmd/ruleset.xml must point at the build-owned duplication gate");
    assertTrue(
        ruleset.contains("build-owned `verifyJavaSemanticShape` gate")
            && ruleset.contains(
                "stricter semantic-shape PMD profile with explicit reviewed exceptions"),
        "gradle/pmd/ruleset.xml must describe how the production semantic-shape profile owns structural PMD rules");
  }

  @Test
  void currentArchitectureGuidanceDoesNotReintroduceTheDeletedProtocolModule() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    String agents = Files.readString(repositoryRoot.resolve("AGENTS.md"));
    String agentsExtra = Files.readString(repositoryRoot.resolve(".codex/AGENTS_EXTRA.md"));
    String developerGuide = Files.readString(repositoryRoot.resolve("docs/DEVELOPER.md"));
    String developerGradle = Files.readString(repositoryRoot.resolve("docs/DEVELOPER_GRADLE.md"));
    String developerJazzer = Files.readString(repositoryRoot.resolve("docs/DEVELOPER_JAZZER.md"));
    String rootConventions =
        Files.readString(
            repositoryRoot.resolve(
                "gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindRootConventionsPlugin.kt"));
    String javaConventions =
        Files.readString(
            repositoryRoot.resolve(
                "gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindJavaConventionsPlugin.kt"));

    assertTrue(
        !agents.contains("- Kotlin 2.4+ / Gradle: " + STALE_KOTLIN_PROTOCOL_LOAD),
        "AGENTS.md must not require the Kotlin protocol for GridGrind work");
    assertTrue(
        agents.contains("GridGrind's Kotlin build logic is not an application surface"),
        "AGENTS.md must teach the Java/Gradle-only rule for GridGrind build logic");
    assertTrue(
        agents.contains(CURRENT_SQLITE_SURFACE),
        "AGENTS.md must publish one current SQLite dependency surface");
    assertTrue(
        !agentsExtra.contains(STALE_MODULE_GRAPH),
        ".codex/AGENTS_EXTRA.md must not teach the deleted cli -> protocol -> engine graph");
    assertTrue(
        agentsExtra.contains("Do not read `AGENTS_KOTLIN24_GRADLE.md`."),
        ".codex/AGENTS_EXTRA.md must preserve the Kotlin-protocol ban");
    assertTrue(
        !agentsExtra.contains(STALE_NULL_DEFAULT_GUIDANCE),
        ".codex/AGENTS_EXTRA.md must not permit null-return protocol default methods");
    assertTrue(
        agentsExtra.contains("`cli -> engine -> contract -> excel-foundation`"),
        ".codex/AGENTS_EXTRA.md must teach the current CLI graph");
    assertTrue(
        agentsExtra.contains("`engine -> excel-foundation`"),
        ".codex/AGENTS_EXTRA.md must teach the engine/foundation boundary");
    assertTrue(
        agentsExtra.contains(
            "Protocol request, response, and discovery DTOs must encode alternative state"),
        ".codex/AGENTS_EXTRA.md must require typed protocol variants instead of null padding");
    assertTrue(
        !developerGuide.contains(STALE_EXECUTOR_RUNTIME_GUIDANCE),
        "docs/DEVELOPER.md must not describe executor as the live runtime bridge");
    assertTrue(
        !developerGuide.contains(STALE_CLI_EXECUTOR_GUIDANCE),
        "docs/DEVELOPER.md must not teach CLI -> executor as the live runtime path");
    assertTrue(
        developerGuide.contains("exported `dev.erst.gridgrind.engine.api` seam"),
        "docs/DEVELOPER.md must name the narrow exported engine API seam");
    assertTrue(
        developerGuide.contains("verifyJavaSourceShape"),
        "docs/DEVELOPER.md must document the build-owned source-shape gate");
    assertTrue(
        developerGuide.contains("verifyControlPlaneShape"),
        "docs/DEVELOPER.md must document the build-owned control-plane shape gate");
    assertTrue(
        developerGuide.contains("verifyJavaSourceDuplication"),
        "docs/DEVELOPER.md must document the build-owned duplication gate");
    assertTrue(
        developerGuide.contains("build/reports/source-shape/java-duplication.tsv"),
        "docs/DEVELOPER.md must publish the duplication report path");
    assertTrue(
        developerGuide.contains("reviewExpiresOn") && developerGuide.contains("splitTrigger"),
        "docs/DEVELOPER.md must teach reviewed exact-surface metadata");
    assertTrue(
        developerGuide.contains("production-safe subset of `design`"),
        "docs/DEVELOPER.md must describe PMD as a scoped design gate rather than the whole category");
    assertTrue(
        developerGuide.contains(
            "engine` coverage also intentionally includes downstream consumer execution data from `executor`"),
        "docs/DEVELOPER.md must explain engine downstream coverage ownership");
    assertTrue(
        developerGradle.contains("verifyJavaSourceShape"),
        "docs/DEVELOPER_GRADLE.md must document the root source-shape task");
    assertTrue(
        developerGradle.contains("verifyControlPlaneShape"),
        "docs/DEVELOPER_GRADLE.md must document the root control-plane shape task");
    assertTrue(
        developerGradle.contains("verifyJavaSourceDuplication"),
        "docs/DEVELOPER_GRADLE.md must document the root duplication task");
    assertTrue(
        developerGradle.contains(
            "Root `check` now depends on `verifyJavaSourceShape`, `verifyControlPlaneShape`"),
        "docs/DEVELOPER_GRADLE.md must teach that source-shape is a build gate");
    assertTrue(
        developerGradle.contains("gradle/build-logic:test"),
        "docs/DEVELOPER_GRADLE.md must teach that build-logic tests are part of the root gate");
    assertTrue(
        developerGradle.contains("verifyNoLegacyBuildSrc"),
        "docs/DEVELOPER_GRADLE.md must teach that legacy buildSrc directories are forbidden");
    assertTrue(
        developerGradle.contains("duplicationGuard")
            && developerGradle.contains("reviewExpiresOn")
            && developerGradle.contains("splitTrigger"),
        "docs/DEVELOPER_GRADLE.md must teach the structural-governance policy metadata");
    assertTrue(
        !developerJazzer.contains(STALE_JAZZER_MODULES),
        "docs/DEVELOPER_JAZZER.md must not claim Jazzer consumes a live protocol module");
    assertTrue(
        developerJazzer.contains(
            "local `engine` and `contract` modules, plus the verification-only `executor` project"),
        "docs/DEVELOPER_JAZZER.md must point at the current local module set");
    assertTrue(
        rootConventions.contains("tasks.register(\"verifyJavaSourceShape\""),
        "GridGrindRootConventionsPlugin must register verifyJavaSourceShape");
    assertTrue(
        rootConventions.contains("tasks.register(\"verifyControlPlaneShape\""),
        "GridGrindRootConventionsPlugin must register verifyControlPlaneShape");
    assertTrue(
        rootConventions.contains("tasks.register(\"verifyJavaSourceDuplication\""),
        "GridGrindRootConventionsPlugin must register verifyJavaSourceDuplication");
    assertTrue(
        rootConventions.contains("checkTask.dependsOn(verifyJavaSourceShape)"),
        "GridGrindRootConventionsPlugin must wire verifyJavaSourceShape into root check");
    assertTrue(
        rootConventions.contains("checkTask.dependsOn(verifyControlPlaneShape)"),
        "GridGrindRootConventionsPlugin must wire verifyControlPlaneShape into root check");
    assertTrue(
        rootConventions.contains("checkTask.dependsOn(verifyJavaSourceDuplication)"),
        "GridGrindRootConventionsPlugin must wire verifyJavaSourceDuplication into root check");
    assertTrue(
        rootConventions.contains("gradle.includedBuild(\"build-logic\").task(\":test\")"),
        "GridGrindRootConventionsPlugin must wire build-logic tests into root check");
    assertTrue(
        javaConventions.contains("tasks.register(\"pmdSemanticMain\", Pmd::class.java)")
            && javaConventions.contains("repositoryLayout.semanticShapePmdRuleset"),
        "GridGrindJavaConventionsPlugin must register the dedicated semantic-shape PMD task");
    assertTrue(
        rootConventions.contains("tasks.register(\"verifyJavaSemanticShape\"")
            && rootConventions.contains(
                "subprojects.map { subproject -> \"${subproject.path}:pmdSemanticMain\" }"),
        "GridGrindRootConventionsPlugin must aggregate pmdSemanticMain into verifyJavaSemanticShape");
  }

  @Test
  void engineModuleExportsRequestExecutionWithoutWorkbookCore() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    String moduleInfo =
        Files.readString(repositoryRoot.resolve("engine/src/main/java/module-info.java"));

    assertTrue(
        QUALIFIED_ENGINE_EXPORT_PATTERN.matcher(moduleInfo).find(),
        "engine module must export the narrow request-execution API package");
    assertTrue(
        !moduleInfo.contains("exports dev.erst.gridgrind.excel;"),
        "engine module must not expose the low-level workbook package");
  }
}
