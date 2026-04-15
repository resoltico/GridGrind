package dev.erst.gridgrind.protocol.parity;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.junit.jupiter.api.Test;

/** Verifies that the public capability inventory stays aligned with the shipped contract. */
final class XlsxParityDocsTest {
  private static final Pattern FRONTMATTER_VERSION_PATTERN =
      Pattern.compile("(?m)^version: \"([^\"]+)\"$");

  @Test
  void publicCapabilityInventoryTracksCurrentReleaseVersion() {
    Path repositoryRoot = repositoryRoot();
    String expectedVersion = releaseVersion(repositoryRoot.resolve("gradle.properties"));

    assertEquals(
        expectedVersion,
        frontmatterVersion(repositoryRoot.resolve("docs/POI_EXCEL_CAPABILITY_INVENTORY.md")));
  }

  @Test
  void obsoleteInternalParityPlanningDocsAreGone() {
    Path repositoryRoot = repositoryRoot();
    assertFalse(
        Files.exists(repositoryRoot.resolve(".codex/POI_5_5_1_XSSF_PARITY_EXECUTION_SPEC.md")));
    assertFalse(Files.exists(repositoryRoot.resolve(".codex/POI_EXCEL_CAPABILITY_INVENTORY.md")));
  }

  @Test
  void publicCapabilityInventoryStaysPublicFacing() {
    String inventory =
        XlsxParitySupport.call(
            "read public capability inventory",
            () ->
                Files.readString(
                    repositoryRoot().resolve("docs/POI_EXCEL_CAPABILITY_INVENTORY.md")));

    assertTrue(inventory.contains("# Apache POI XSSF `.xlsx` Capability Inventory"));
    assertTrue(inventory.contains("## Capability Matrix"));
    assertFalse(inventory.contains("Phase "));
    assertFalse(inventory.contains("milestone"));
    assertFalse(inventory.contains(".codex/"));
    assertFalse(inventory.contains("OUTSIDE_CURRENT_PARITY_TARGET"));
    assertFalse(inventory.contains("canonical parity ledger"));
  }

  @Test
  void publicCapabilityInventoryCapturesCurrentContractBoundaries() {
    String inventory =
        XlsxParitySupport.call(
            "read public capability inventory",
            () ->
                Files.readString(
                    repositoryRoot().resolve("docs/POI_EXCEL_CAPABILITY_INVENTORY.md")));

    assertTrue(inventory.contains("| Sheet copy | Documented API support | `PARTIAL` |"));
    assertTrue(
        inventory.contains(
            "| Sheet presentation and sheet-level defaults | Documented XSSF sheet-view and worksheet APIs | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "screen display flags, right-to-left layout, tab color, outline-summary placement,"));
    assertTrue(
        inventory.contains(
            "| Threaded comments | Specialized XSSF support beyond classic comments | `NOT_EXPOSED` |"));
    assertTrue(
        inventory.contains(
            "| Array-formula authoring and array-group metadata | XSSF array-formula APIs | `NOT_EXPOSED` |"));
    assertTrue(
        inventory.contains(
            "| Formula surface summaries | GridGrind-derived read over POI formula facts | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Formula health analysis | GridGrind-derived analysis over POI formula facts | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Hyperlink health analysis | GridGrind-derived analysis over POI hyperlink facts | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Named-range surface summaries | GridGrind-derived read over POI named-range facts | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Named-range health analysis | GridGrind-derived analysis over POI named-range facts | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Sparkline discovery and authoring | XSSF sheet sparkline APIs | `NOT_EXPOSED` |"));
    assertTrue(
        inventory.contains(
            "| OOXML encryption, password-protected package open or save, and XML signing |"));
    assertTrue(
        inventory.contains(
            "| XML-mapped tables and slicers | Broader XSSF table ecosystem | `NOT_EXPOSED` |"));
    assertTrue(
        inventory.contains(
            "| Sheet schema inference | GridGrind-derived read over cell snapshots | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Aggregate workbook findings | GridGrind-derived aggregate over all shipped health families | `COMPLETE` |"));
    assertTrue(inventory.contains("print-gridline output"));
    assertTrue(inventory.contains("Currently not exposed:"));
    assertFalse(
        inventory.contains(
            "| Formula surface, sheet schema, named-range surface, and aggregate workbook findings |"));
    assertFalse(inventory.contains("ExcelCommentSupport.java"));
    assertTrue(inventory.contains("engine/excel/ExcelComment.java"));
    assertTrue(inventory.contains("engine/excel/ExcelCommentAnchor.java"));
    assertTrue(inventory.contains("FORCE_FORMULA_RECALCULATION_ON_OPEN"));
    assertFalse(inventory.contains("FORCE_FORMULA_RECALC_ON_OPEN"));
    assertTrue(inventory.contains("`LAMBDA` and `LET` are currently rejected"));
    assertTrue(inventory.contains("analysis.totalFormulaCellCount"));
    assertTrue(inventory.contains("analysis.checkedNamedRangeCount"));
    assertTrue(inventory.contains("- drawing-family sheet copy"));
  }

  @Test
  void publicStreamingWriteDocsUseCanonicalRecalcOnOpenOperationName() {
    Path repositoryRoot = repositoryRoot();
    for (String relativePath :
        List.of(
            "README.md",
            "docs/OPERATIONS.md",
            "docs/QUICK_REFERENCE.md",
            "docs/LIMITATIONS.md",
            "docs/POI_EXCEL_CAPABILITY_INVENTORY.md")) {
      String document =
          XlsxParitySupport.call(
              "read " + relativePath, () -> Files.readString(repositoryRoot.resolve(relativePath)));

      assertTrue(
          document.contains("FORCE_FORMULA_RECALCULATION_ON_OPEN"),
          relativePath + " must use the canonical streaming-write operation name");
      assertFalse(
          document.contains("FORCE_FORMULA_RECALC_ON_OPEN"),
          relativePath + " must not use the rejected legacy shorthand");
    }
  }

  @Test
  void publicDocsDescribeCurrentFormulaAndAnalysisReadContracts() {
    Path repositoryRoot = repositoryRoot();
    for (String relativePath :
        List.of("README.md", "docs/OPERATIONS.md", "docs/QUICK_REFERENCE.md", "docs/ERRORS.md")) {
      String document =
          XlsxParitySupport.call(
              "read " + relativePath, () -> Files.readString(repositoryRoot.resolve(relativePath)));

      assertTrue(
          document.contains("LAMBDA") && document.contains("LET"),
          relativePath + " must mention the current LAMBDA/LET limitation");
      assertTrue(
          document.contains("INVALID_FORMULA"),
          relativePath + " must describe the INVALID_FORMULA boundary");
    }

    String operations =
        XlsxParitySupport.call(
            "read docs/OPERATIONS.md",
            () -> Files.readString(repositoryRoot.resolve("docs/OPERATIONS.md")));
    assertTrue(operations.contains("totalFormulaCellCount"));
    assertTrue(operations.contains("checkedNamedRangeCount"));

    String quickReference =
        XlsxParitySupport.call(
            "read docs/QUICK_REFERENCE.md",
            () -> Files.readString(repositoryRoot.resolve("docs/QUICK_REFERENCE.md")));
    assertTrue(quickReference.contains("analysis.summary"));
  }

  private static Path repositoryRoot() {
    Path candidate = Path.of("").toAbsolutePath().normalize();
    while (candidate != null) {
      if (Files.exists(candidate.resolve("gradle.properties"))
          && Files.exists(candidate.resolve("docs/POI_EXCEL_CAPABILITY_INVENTORY.md"))) {
        return candidate;
      }
      candidate = candidate.getParent();
    }
    throw new AssertionError("Could not locate the GridGrind repository root.");
  }

  private static String releaseVersion(Path gradlePropertiesPath) {
    List<String> lines =
        XlsxParitySupport.call(
            "read gradle.properties", () -> Files.readAllLines(gradlePropertiesPath));
    return lines.stream()
        .filter(line -> line.startsWith("version="))
        .findFirst()
        .map(line -> line.substring("version=".length()))
        .orElseThrow(() -> new AssertionError("No version= entry found in gradle.properties"));
  }

  private static String frontmatterVersion(Path documentPath) {
    String document =
        XlsxParitySupport.call(
            "read " + documentPath.getFileName(), () -> Files.readString(documentPath));
    Matcher matcher = FRONTMATTER_VERSION_PATTERN.matcher(document);
    assertTrue(matcher.find(), "Missing frontmatter version in " + documentPath);
    return matcher.group(1);
  }
}
