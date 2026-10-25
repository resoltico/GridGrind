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
  private static final String STALE_JAZZER_MODULES = "local `engine` and `protocol` modules";
  private static final String STALE_EXECUTOR_RUNTIME_GUIDANCE =
      "executor/   The only execution bridge from the canonical contract into the";
  private static final String STALE_CLI_EXECUTOR_GUIDANCE =
      "or --request file, delegates to executor, writes the response,";

  @Test
  void currentArchitectureGuidanceDoesNotReintroduceTheDeletedProtocolModule() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    String developerGuide = Files.readString(repositoryRoot.resolve("docs/DEVELOPER.md"));
    String developerGradle = Files.readString(repositoryRoot.resolve("docs/DEVELOPER_GRADLE.md"));
    String developerJazzer = Files.readString(repositoryRoot.resolve("docs/DEVELOPER_JAZZER.md"));
    String rootConventions =
        Files.readString(
            repositoryRoot.resolve(
                "gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindRootConventionsPlugin.kt"));

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
        developerGuide.contains("production-safe subset of `design`"),
        "docs/DEVELOPER.md must describe PMD as a scoped design gate rather than the whole category");
    assertTrue(
        developerGuide.contains(
            "engine` coverage also intentionally includes downstream consumer execution data from `executor`"),
        "docs/DEVELOPER.md must explain engine downstream coverage ownership");
    assertTrue(
        developerGradle.contains("gradle/build-logic:test"),
        "docs/DEVELOPER_GRADLE.md must teach that build-logic tests are part of the root gate");
    assertTrue(
        developerGradle.contains("verifyNoLegacyBuildSrc"),
        "docs/DEVELOPER_GRADLE.md must teach that legacy buildSrc directories are forbidden");
    assertTrue(
        !developerJazzer.contains(STALE_JAZZER_MODULES),
        "docs/DEVELOPER_JAZZER.md must not claim Jazzer consumes a live protocol module");
    assertTrue(
        developerJazzer.contains(
            "local `engine` and `contract` modules, plus the verification-only `executor` project"),
        "docs/DEVELOPER_JAZZER.md must point at the current local module set");
    assertTrue(
        rootConventions.contains("gradle.includedBuild(\"build-logic\").task(\":test\")"),
        "GridGrindRootConventionsPlugin must wire build-logic tests into root check");
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
