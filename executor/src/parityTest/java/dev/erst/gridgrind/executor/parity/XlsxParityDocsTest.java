package dev.erst.gridgrind.executor.parity;

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
  private static final List<String> SPLIT_REFERENCE_DOCS =
      List.of(
          "docs/REQUEST_AND_EXECUTION_REFERENCE.md",
          "docs/WORKBOOK_AND_LAYOUT_MUTATIONS.md",
          "docs/CELL_AND_DRAWING_MUTATIONS.md",
          "docs/STRUCTURED_FEATURE_MUTATIONS.md",
          "docs/ASSERTION_AND_INSPECTION_REFERENCE.md");
  private static final List<String> PUBLIC_REFERENCE_DOCS =
      List.of(
          "docs/OPERATIONS.md",
          "docs/QUICK_REFERENCE.md",
          "docs/REQUEST_AND_EXECUTION_REFERENCE.md",
          "docs/WORKBOOK_AND_LAYOUT_MUTATIONS.md",
          "docs/CELL_AND_DRAWING_MUTATIONS.md",
          "docs/STRUCTURED_FEATURE_MUTATIONS.md",
          "docs/ASSERTION_AND_INSPECTION_REFERENCE.md");

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

    assertTrue(inventory.contains("| Sheet copy | Documented API support | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Sheet presentation and sheet-level defaults | Documented XSSF sheet-view and worksheet APIs | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "screen display flags, right-to-left layout, tab color, outline-summary placement,"));
    assertTrue(
        inventory.contains(
            "| Threaded comments | OOXML threaded-comment schemas exist, but no first-class XSSF threaded-comment usermodel APIs were found in audited POI 5.5.1 docs/source | `NOT_EXPOSED` |"));
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
            "| Sparkline discovery and authoring | No first-class sparkline XSSF usermodel APIs were found in audited POI 5.5.1 docs/source | `NOT_EXPOSED` |"));
    assertTrue(
        inventory.contains(
            "| OOXML encryption, password-protected package open or save, and XML signing |"));
    assertTrue(
        inventory.contains(
            "| Array-formula authoring and array-group metadata | `Sheet.setArrayFormula(...)`, `Sheet.removeArrayFormula(...)`, and cell array-group metadata APIs | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Custom XML mappings and XML import/export | `XSSFWorkbook.getCustomXMLMappings()`, `getMapInfo()`, `XSSFImportFromXML`, and `XSSFExportToXml` | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Slicers | No first-class XSSF slicer usermodel APIs were found in audited POI 5.5.1 docs/source | `NOT_EXPOSED` |"));
    assertTrue(
        inventory.contains(
            "| Broader XDDF chart authoring families and combo charts | XDDF `ChartTypes`/`createData(...)` support plus official scatter and combo examples | `COMPLETE` |"));
    assertTrue(
        inventory.contains(
            "| Signature-line OOXML metadata | `XSSFSignatureLine` and `SignatureLine` APIs | `COMPLETE` |"));
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
    assertTrue(inventory.contains("markRecalculateOnOpen"));
    assertTrue(inventory.contains("EVALUATE_ALL"));
    assertTrue(inventory.contains("CLEAR_CACHES_ONLY"));
    assertFalse(inventory.contains("FORCE_FORMULA_RECALCULATION_ON_OPEN"));
    assertFalse(inventory.contains("FORCE_FORMULA_RECALC_ON_OPEN"));
    assertTrue(inventory.contains("`LAMBDA` and `LET` are currently rejected"));
    assertTrue(inventory.contains("analysis.totalFormulaCellCount"));
    assertTrue(inventory.contains("analysis.checkedNamedRangeCount"));
    assertFalse(inventory.contains("- drawing-family sheet copy"));
    assertFalse(inventory.contains("- signature lines"));
  }

  @Test
  void publicStreamingWriteDocsUseCalculationPolicyTerminology() {
    Path repositoryRoot = repositoryRoot();
    for (String relativePath :
        List.of(
            "README.md",
            "docs/OPERATIONS.md",
            "docs/QUICK_REFERENCE.md",
            "docs/REQUEST_AND_EXECUTION_REFERENCE.md",
            "docs/STRUCTURED_FEATURE_MUTATIONS.md",
            "docs/LIMITATIONS.md",
            "docs/POI_EXCEL_CAPABILITY_INVENTORY.md")) {
      String document =
          XlsxParitySupport.call(
              "read " + relativePath, () -> Files.readString(repositoryRoot.resolve(relativePath)));

      assertTrue(
          document.contains("markRecalculateOnOpen"),
          relativePath + " must explain open-time recalculation through execution.calculation");
      assertTrue(
          document.contains("ENSURE_SHEET") && document.contains("APPEND_ROW"),
          relativePath + " must describe the remaining STREAMING_WRITE mutation actions");
      assertFalse(
          document.contains("FORCE_FORMULA_RECALCULATION_ON_OPEN"),
          relativePath + " must not refer to the deleted recalc mutation action");
      assertFalse(
          document.contains("FORCE_FORMULA_RECALC_ON_OPEN"),
          relativePath + " must not use the rejected legacy shorthand");
    }
  }

  @Test
  void publicReferenceDocsStaySplitIntoStableEntryPoints() {
    Path repositoryRoot = repositoryRoot();
    String operations = readDoc(repositoryRoot, "docs/OPERATIONS.md");
    String quickReference = readDoc(repositoryRoot, "docs/QUICK_REFERENCE.md");

    for (String relativePath : SPLIT_REFERENCE_DOCS) {
      assertTrue(
          Files.exists(repositoryRoot.resolve(relativePath)),
          relativePath + " must exist as part of the split public reference surface");
      assertTrue(
          operations.contains(Path.of(relativePath).getFileName().toString()),
          "docs/OPERATIONS.md must link to " + relativePath);
      assertTrue(
          quickReference.contains(Path.of(relativePath).getFileName().toString()),
          "docs/QUICK_REFERENCE.md must link to " + relativePath);
    }

    assertTrue(
        lineCount(operations) <= 220,
        "docs/OPERATIONS.md must stay an index instead of growing back into a god-file");
    assertTrue(
        lineCount(quickReference) <= 400,
        "docs/QUICK_REFERENCE.md must stay a cheat sheet instead of regrowing into a full manual");
  }

  @Test
  void publicDocsDescribeCurrentFormulaAndAnalysisReadContracts() {
    Path repositoryRoot = repositoryRoot();
    String formulaAndAnalysisCorpus =
        readAllDocs(repositoryRoot, PUBLIC_REFERENCE_DOCS)
            + readDoc(repositoryRoot, "README.md")
            + readDoc(repositoryRoot, "docs/ERRORS.md");

    assertTrue(
        formulaAndAnalysisCorpus.contains("LAMBDA") && formulaAndAnalysisCorpus.contains("LET"),
        "public docs must mention the current LAMBDA/LET limitation");
    assertTrue(
        formulaAndAnalysisCorpus.contains("INVALID_FORMULA"),
        "public docs must describe the INVALID_FORMULA boundary");
    assertTrue(
        formulaAndAnalysisCorpus.contains("analysis.totalFormulaCellCount"),
        "public docs must describe formula-surface aggregate counts");
    assertTrue(
        formulaAndAnalysisCorpus.contains("analysis.checkedNamedRangeCount"),
        "public docs must describe named-range-health aggregate counts");
    assertTrue(
        formulaAndAnalysisCorpus.contains("analysis.summary"),
        "public docs must describe analysis summary payloads");
  }

  @Test
  void publicDocsDescribeCurrentChartAndSignatureLineContracts() {
    Path repositoryRoot = repositoryRoot();
    String operations = readDoc(repositoryRoot, "docs/OPERATIONS.md");
    String quickReference = readDoc(repositoryRoot, "docs/QUICK_REFERENCE.md");
    String cellAndDrawing = readDoc(repositoryRoot, "docs/CELL_AND_DRAWING_MUTATIONS.md");
    String assertionsAndInspections =
        readDoc(repositoryRoot, "docs/ASSERTION_AND_INSPECTION_REFERENCE.md");
    String readme = readDoc(repositoryRoot, "README.md");
    String limitations = readDoc(repositoryRoot, "docs/LIMITATIONS.md");

    assertTrue(operations.contains("CELL_AND_DRAWING_MUTATIONS.md"));
    assertTrue(cellAndDrawing.contains("### SET_SIGNATURE_LINE"));
    assertTrue(cellAndDrawing.contains("\"plots\": ["));
    assertTrue(cellAndDrawing.contains("SIGNATURE_LINE"));
    assertTrue(assertionsAndInspections.contains("SIGNATURE_LINE"));
    assertTrue(cellAndDrawing.contains("`GIF`, `TIFF`, `EPS`, `BMP`,"));
    assertTrue(cellAndDrawing.contains("or `WPG`."));
    assertFalse(
        cellAndDrawing.contains("Supported authored families are `BAR`, `LINE`, and `PIE`."));

    assertTrue(quickReference.contains("SET_SIGNATURE_LINE"));
    assertTrue(
        quickReference.contains(
            "Supported authored plot families are `AREA`, `AREA_3D`, `BAR`, `BAR_3D`, `DOUGHNUT`, `LINE`,"));
    assertFalse(
        quickReference.contains("Supported authored families are `BAR`, `LINE`, and `PIE`."));

    assertTrue(readme.contains("examples/signature-line-request.json"));
    assertTrue(limitations.contains("AREA_3D"));
    assertTrue(limitations.contains("SURFACE_3D"));
    assertTrue(limitations.contains("UNSUPPORTED"));
    assertFalse(
        limitations.contains("Supported for factual reads and authored `BAR`, `LINE`, and `PIE`"));
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

  private static String readDoc(Path repositoryRoot, String relativePath) {
    return XlsxParitySupport.call(
        "read " + relativePath, () -> Files.readString(repositoryRoot.resolve(relativePath)));
  }

  private static String readAllDocs(Path repositoryRoot, List<String> relativePaths) {
    StringBuilder builder = new StringBuilder();
    for (String relativePath : relativePaths) {
      builder.append(readDoc(repositoryRoot, relativePath)).append('\n');
    }
    return builder.toString();
  }

  private static long lineCount(String document) {
    return document.lines().count();
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
