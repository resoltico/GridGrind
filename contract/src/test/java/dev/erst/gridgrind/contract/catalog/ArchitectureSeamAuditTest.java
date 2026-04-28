package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiPredicate;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Build-failing architectural audits over key module and seam boundaries. */
class ArchitectureSeamAuditTest {
  private static final Pattern CONTRACT_ENGINE_IMPORT_PATTERN =
      Pattern.compile(
          "^import dev\\.erst\\.gridgrind\\.excel\\.(?!foundation\\.)", Pattern.MULTILINE);
  private static final Pattern POI_PRIVATE_ACCESS_PATTERN =
      Pattern.compile("privateLookupIn|getDeclaredField|getDeclaredMethod|setAccessible\\(");
  private static final String STALE_MODULE_GRAPH = "`cli -> protocol -> engine`";
  private static final String STALE_JAZZER_MODULES = "local `engine` and `protocol` modules";
  private static final String STALE_NULL_DEFAULT_GUIDANCE = "The sole sanctioned null-return site";
  private static final String STALE_PMD_PROTOCOL_ORCHESTRATOR = "GridGrindService";
  private static final String STALE_PMD_EXCEPTION_HELPER = "withExceptionData()";

  @Test
  void contractModuleOnlyImportsExcelFoundationTypes() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    List<String> violations =
        matchingFiles(
            repositoryRoot.resolve("contract/src/main/java"),
            (path, contents) -> CONTRACT_ENGINE_IMPORT_PATTERN.matcher(contents).find());

    assertTrue(
        violations.isEmpty(),
        () -> "contract module must not import engine-side excel types: " + violations);
  }

  @Test
  void formulaWritesStayCentralized() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    Path engineSourceRoot = repositoryRoot.resolve("engine/src/main/java/dev/erst/gridgrind/excel");
    List<String> violations =
        matchingFiles(
            engineSourceRoot,
            (path, contents) ->
                contents.contains("setCellFormula(")
                    && !"ExcelFormulaWriteSupport.java".equals(path.getFileName().toString()));

    assertTrue(
        violations.isEmpty(),
        () ->
            "Direct formula writes must stay centralized in ExcelFormulaWriteSupport: "
                + violations);
  }

  @Test
  void poiPrivateAccessStaysCentralized() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    Path engineSourceRoot = repositoryRoot.resolve("engine/src/main/java/dev/erst/gridgrind/excel");
    List<String> violations =
        matchingFiles(
            engineSourceRoot,
            (path, contents) ->
                POI_PRIVATE_ACCESS_PATTERN.matcher(contents).find()
                    && !"PoiPrivateAccessSupport.java".equals(path.getFileName().toString()));

    assertTrue(
        violations.isEmpty(),
        () ->
            "POI reflective/private access must stay centralized in PoiPrivateAccessSupport: "
                + violations);
  }

  @Test
  void keyGodFilesStaySplit() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();

    assertLineCountAtMost(
        repositoryRoot.resolve(
            "contract/src/main/java/dev/erst/gridgrind/contract/catalog/GridGrindProtocolCatalog.java"),
        600);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "contract/src/main/java/dev/erst/gridgrind/contract/catalog/GridGrindProtocolCatalogFieldGroupSupport.java"),
        100);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "contract/src/main/java/dev/erst/gridgrind/contract/catalog/GridGrindProtocolCatalogNestedTypeGroups.java"),
        950);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "contract/src/main/java/dev/erst/gridgrind/contract/catalog/GridGrindProtocolCatalogStyleTypeGroups.java"),
        200);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "contract/src/main/java/dev/erst/gridgrind/contract/dto/GridGrindResponse.java"),
        875);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "contract/src/main/java/dev/erst/gridgrind/contract/dto/ProblemContext.java"),
        998);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "engine/src/main/java/dev/erst/gridgrind/excel/WorkbookCommand.java"),
        975);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "executor/src/main/java/dev/erst/gridgrind/executor/WorkbookCommandStructuredInputConverter.java"),
        875);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "jazzer/src/main/java/dev/erst/gridgrind/jazzer/support/OperationSequenceModel.java"),
        150);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "jazzer/src/main/java/dev/erst/gridgrind/jazzer/support/OperationSequenceValueFactory.java"),
        925);
  }

  @Test
  void productionJavaFilesStayBelowTheAbsoluteCeiling() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    List<String> violations = new ArrayList<>();
    List<Path> sourceRoots =
        List.of(
            repositoryRoot.resolve("authoring-java/src/main/java"),
            repositoryRoot.resolve("cli/src/main/java"),
            repositoryRoot.resolve("contract/src/main/java"),
            repositoryRoot.resolve("engine/src/main/java"),
            repositoryRoot.resolve("excel-foundation/src/main/java"),
            repositoryRoot.resolve("executor/src/main/java"),
            repositoryRoot.resolve("jazzer/src/main/java"));
    for (Path sourceRoot : sourceRoots) {
      if (!Files.isDirectory(sourceRoot)) {
        continue;
      }
      try (Stream<Path> files = Files.walk(sourceRoot)) {
        for (Path path :
            files
                .filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".java"))
                .toList()) {
          long lineCount;
          try (Stream<String> lines = Files.lines(path)) {
            lineCount = lines.count();
          }
          if (lineCount > 1000L) {
            violations.add(path + " (" + lineCount + " lines)");
          }
        }
      }
    }
    assertTrue(
        violations.isEmpty(),
        () -> "Production Java sources must stay under 1000 lines: " + violations);
  }

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
        ruleset.contains("ArchitectureSeamAuditTest"),
        "gradle/pmd/ruleset.xml must point at the active split ratchet owner");
  }

  @Test
  void currentArchitectureGuidanceDoesNotReintroduceTheDeletedProtocolModule() throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    String agentsExtra = Files.readString(repositoryRoot.resolve(".codex/AGENTS_EXTRA.md"));
    String developerJazzer = Files.readString(repositoryRoot.resolve("docs/DEVELOPER_JAZZER.md"));

    assertTrue(
        !agentsExtra.contains(STALE_MODULE_GRAPH),
        ".codex/AGENTS_EXTRA.md must not teach the deleted cli -> protocol -> engine graph");
    assertTrue(
        !agentsExtra.contains(STALE_NULL_DEFAULT_GUIDANCE),
        ".codex/AGENTS_EXTRA.md must not permit null-return protocol default methods");
    assertTrue(
        agentsExtra.contains("`cli -> executor -> contract -> excel-foundation`"),
        ".codex/AGENTS_EXTRA.md must teach the current CLI graph");
    assertTrue(
        agentsExtra.contains("`executor -> engine -> excel-foundation`"),
        ".codex/AGENTS_EXTRA.md must teach the executor/engine boundary");
    assertTrue(
        agentsExtra.contains(
            "Protocol request, response, and discovery DTOs must encode alternative state"),
        ".codex/AGENTS_EXTRA.md must require typed protocol variants instead of null padding");
    assertTrue(
        !developerJazzer.contains(STALE_JAZZER_MODULES),
        "docs/DEVELOPER_JAZZER.md must not claim Jazzer consumes a live protocol module");
    assertTrue(
        developerJazzer.contains("local `engine`, `contract`, and `executor` modules"),
        "docs/DEVELOPER_JAZZER.md must point at the current local module set");
  }

  private static List<String> matchingFiles(Path root, BiPredicate<Path, String> matcher)
      throws IOException {
    List<String> violations = new ArrayList<>();
    try (Stream<Path> files = Files.walk(root)) {
      for (Path path :
          files
              .filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".java"))
              .toList()) {
        String contents = Files.readString(path);
        if (matcher.test(path, contents)) {
          violations.add(root.getParent().getParent().getParent().relativize(path).toString());
        }
      }
    }
    return violations;
  }

  private static void assertLineCountAtMost(Path path, int limit) throws IOException {
    long lineCount;
    try (Stream<String> lines = Files.lines(path)) {
      lineCount = lines.count();
    }
    assertTrue(
        lineCount <= limit,
        () -> path + " must stay under " + limit + " lines but is " + lineCount);
  }
}
