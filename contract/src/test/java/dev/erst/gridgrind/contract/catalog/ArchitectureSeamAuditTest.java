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
        1800);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "contract/src/main/java/dev/erst/gridgrind/contract/catalog/GridGrindProtocolCatalogFieldGroupSupport.java"),
        2000);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "jazzer/src/main/java/dev/erst/gridgrind/jazzer/support/OperationSequenceModel.java"),
        1400);
    assertLineCountAtMost(
        repositoryRoot.resolve(
            "jazzer/src/main/java/dev/erst/gridgrind/jazzer/support/OperationSequenceValueFactory.java"),
        1500);
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
