package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Build-failing audit over public XLSX capability and limitation docs. */
class GridGrindSpreadsheetDocsAuditTest {
  @Test
  void poiCapabilityInventoryReflectsAuditedUnexposedFamilies() throws IOException {
    String inventory = readDoc("docs/POI_EXCEL_CAPABILITY_INVENTORY.md");

    assertAll(
        () ->
            assertTrue(
                inventory.contains("| Custom XML mappings and XML import/export |"),
                "inventory must list unexposed custom XML mapping support"),
        () ->
            assertTrue(
                inventory.contains("| Slicers |"),
                "inventory must list slicers separately from custom XML mappings"),
        () ->
            assertTrue(
                inventory.contains("| Signature-line OOXML metadata |"),
                "inventory must list signature-line support separately from package signing"),
        () ->
            assertTrue(
                inventory.contains("| Broader XDDF chart authoring families and combo charts |"),
                "inventory must list non-productized XDDF chart authoring families"),
        () ->
            assertFalse(
                inventory.contains("| XML-mapped tables and slicers |"),
                "inventory must not conflate custom XML mappings with slicers"),
        () ->
            assertFalse(
                inventory.contains(
                    "| Sparkline discovery and authoring | XSSF sheet sparkline APIs |"),
                "inventory must not overstate sparkline usermodel support"),
        () ->
            assertFalse(
                inventory.contains(
                    "even though Apache POI exposes limited sparkline APIs at the XSSF sheet layer"),
                "inventory must not claim audited sparkline support that was not found"),
        () ->
            assertFalse(
                inventory.contains(
                    "| Threaded comments | Specialized XSSF support beyond classic comments |"),
                "inventory must not overstate threaded-comment usermodel support"));
  }

  @Test
  void poiCapabilityInventoryKeepsPublicSurfaceEvidenceAndPrimarySources() throws IOException {
    String inventory = readDoc("docs/POI_EXCEL_CAPABILITY_INVENTORY.md");

    List<String> shippedRowsMissingContractEvidence =
        capabilityRows(inventory).stream()
            .filter(CapabilityRow::shipped)
            .filter(row -> !row.evidence().contains("contract/"))
            .map(CapabilityRow::capability)
            .toList();

    assertAll(
        () ->
            assertTrue(
                shippedRowsMissingContractEvidence.isEmpty(),
                () ->
                    "every COMPLETE/PARTIAL inventory row must cite public contract evidence: "
                        + shippedRowsMissingContractEvidence),
        () ->
            assertTrue(
                inventory.contains("[HSSF and XSSF Examples]"),
                "inventory must cite the official POI examples page"),
        () ->
            assertTrue(
                inventory.contains("[REL_5_5_1 `XSSFSheet.java`]"),
                "inventory must cite the audited XSSFSheet source"),
        () ->
            assertTrue(
                inventory.contains("[REL_5_5_1 `XSSFSignatureLine.java`]"),
                "inventory must cite the audited XSSFSignatureLine source"),
        () ->
            assertTrue(
                inventory.contains("[REL_5_5_1 `MapInfo.java`]"),
                "inventory must cite the audited MapInfo source"),
        () ->
            assertTrue(
                inventory.contains("[REL_5_5_1 `XSSFImportFromXML.java`]"),
                "inventory must cite the audited XML import source"),
        () ->
            assertTrue(
                inventory.contains("[REL_5_5_1 `XSSFExportToXml.java`]"),
                "inventory must cite the audited XML export source"));
  }

  @Test
  void limitationsRegistryStatesEnforcementAndFormatScopeTruthfully() throws IOException {
    String limitations = readDoc("docs/LIMITATIONS.md");
    String normalizedLimitations = limitations.replaceAll("\\s+", " ");

    assertAll(
        () ->
            assertFalse(
                limitations.contains("GridGrind does not currently enforce these at request time"),
                "limitations doc must not claim blanket non-enforcement for Excel ceilings"),
        () ->
            assertTrue(
                normalizedLimitations.contains(
                    "Some are already enforced on relevant public request paths"),
                "limitations doc must explain that some Excel ceilings are preflighted"),
        () ->
            assertTrue(
                normalizedLimitations.contains(
                    "Product-enforced limits are expected to stay traceable into validation code"),
                "limitations doc must distinguish product-enforced traceability from doc-only ceilings"),
        () ->
            assertFalse(
                normalizedLimitations.contains(
                    "IDs are referenced in the validation code that enforces the limit, in catalog summaries, and in help text"),
                "limitations doc must not claim blanket code/help/catalog propagation for every LIM id"),
        () ->
            assertFalse(
                limitations.contains(
                    "These are hard ceilings of the `.xlsx` format. They are reflected in\n`SpreadsheetVersion.EXCEL2007` in Apache POI 5.5.1."),
                "limitations doc must not claim every structural limit comes from SpreadsheetVersion"),
        () ->
            assertTrue(
                normalizedLimitations.contains(
                    "Many, such as max rows, max columns, max text length, max styles, and max"),
                "limitations doc must distinguish SpreadsheetVersion-backed ceilings from other limits"),
        () ->
            assertTrue(
                normalizedLimitations.contains(
                    "Where POI 5.5.1 does not expose a dedicated constant (for example, hyperlink count, formula"),
                "limitations doc must explain when it is citing Excel-published limits directly"),
        () ->
            assertTrue(
                limitations.contains(
                    "| **Category** | GridGrind (enforces Excel limit for addressed row indices and bands) |"),
                "LIM-008 must declare request-path enforcement"),
        () ->
            assertTrue(
                limitations.contains(
                    "| **Category** | GridGrind (enforces Excel limit for addressed column indices and bands) |"),
                "LIM-009 must declare request-path enforcement"),
        () ->
            assertFalse(
                limitations.contains(
                    "GridGrind uses Apache POI XSSF, which implements only the `.xlsx` (OOXML) format."),
                "LIM-002 must not claim that upstream XSSF only handles plain .xlsx"),
        () ->
            assertTrue(
                limitations.contains(
                    "GridGrind intentionally narrows Apache POI's broader OOXML workbook support to plain `.xlsx`."),
                "LIM-002 must state that the .xlsx-only restriction is product-owned"),
        () ->
            assertFalse(
                limitations.contains(
                    "| Macros (VBA/XLM) | Read: preserved. Write: not creatable. |"),
                "unsupported-features table must not describe .xlsm handling as if it were in scope"),
        () ->
            assertTrue(
                limitations.contains(
                    "| Macro-enabled OOXML (`.xlsm`) | Out of scope for the shipped GridGrind contract; LIM-002 rejects it even though Apache POI can preserve and extract VBA from `.xlsm` packages. |"),
                "unsupported-features table must describe .xlsm as an out-of-scope product choice"),
        () ->
            assertFalse(
                limitations.contains("WorkbookOperation.Validation"),
                "limitations doc must point at current validation owners"),
        () ->
            assertTrue(
                limitations.contains("### LIM-022"),
                "limitations doc must register the zoom ceiling"),
        () ->
            assertTrue(
                limitations.contains("409.0"),
                "limitations doc must use the Excel row-height ceiling"),
        () ->
            assertTrue(
                limitations.contains("| **Category** | Excel format |"),
                "limitations doc must label Excel-only ceilings without implying a POI constant"),
        () ->
            assertTrue(
                limitations.contains("quick-guide.html"),
                "limitations doc must cite the official POI quick guide"),
        () ->
            assertTrue(
                limitations.contains(
                    "REL_5_5_1/poi-ooxml/src/main/java/org/apache/poi/xssf/usermodel/XSSFSheet.java"),
                "limitations doc must cite the audited XSSFSheet source"),
        () ->
            assertTrue(
                limitations.contains("MutationAction.Validation.requireZoomPercent"),
                "zoom limit entry must point at the current validation path"));
  }

  @Test
  void quickStartUsesArtifactNativeBudgetBootstrapFlow() throws IOException {
    String quickStart = readDoc("docs/QUICK_START.md");

    assertAll(
        () ->
            assertTrue(
                quickStart.contains("--print-example BUDGET"),
                "quick start must teach the built-in budget example bootstrap flow"),
        () ->
            assertFalse(
                quickStart.contains(
                    "copy `examples/budget-request.json` into your working directory first"),
                "quick start must not require a repo checkout for first-run artifact usage"));
  }

  @Test
  void examplesGuideDocumentsBuiltInExamplePortabilityTruthfully() throws IOException {
    String examples = readDoc("docs/EXAMPLES.md");
    String normalizedExamples = examples.replaceAll("\\s+", " ");

    assertAll(
        () ->
            assertTrue(
                normalizedExamples.contains(
                    "Self-contained built-ins can run from a blank artifact workspace"),
                "examples guide must describe self-contained built-ins"),
        () ->
            assertTrue(
                normalizedExamples.contains(
                    "Repo-asset-backed built-ins require the matching `examples/` asset directories"),
                "examples guide must describe repo-asset-backed built-ins honestly"),
        () ->
            assertTrue(
                examples.contains("CUSTOM_XML")
                    && examples.contains("SOURCE_BACKED_INPUT")
                    && examples.contains("PACKAGE_SECURITY_INSPECTION"),
                "examples guide must name the repo-asset-backed built-ins"),
        () ->
            assertTrue(
                normalizedExamples.contains("shippedExamples[*].workspaceMode")
                    && normalizedExamples.contains("shippedExamples[*].requiredPaths"),
                "examples guide must document the machine-readable prerequisite metadata"));
  }

  @Test
  void docsDescribeDoctorScopeTruthfully() throws IOException {
    String requestReference = readDoc("docs/REQUEST_AND_EXECUTION_REFERENCE.md");
    String quickStart = readDoc("docs/QUICK_START.md");
    String quickReference = readDoc("docs/QUICK_REFERENCE.md");
    String operations = readDoc("docs/OPERATIONS.md");
    String normalizedRequestReference = requestReference.replaceAll("\\s+", " ");
    String normalizedQuickStart = quickStart.replaceAll("\\s+", " ");
    String normalizedQuickReference = quickReference.replaceAll("\\s+", " ");
    String normalizedOperations = operations.replaceAll("\\s+", " ");

    assertAll(
        () ->
            assertTrue(
                normalizedRequestReference.contains("source-backed authored input resolution"),
                "request reference must explain that doctor preflights source-backed inputs"),
        () ->
            assertTrue(
                normalizedRequestReference.contains(
                    "existing workbook-source accessibility without mutating a workbook"),
                "request reference must explain the doctor workbook-source preflight boundary"),
        () ->
            assertTrue(
                normalizedRequestReference.contains("OPEN_WORKBOOK"),
                "request reference must explain that doctor can fail during OPEN_WORKBOOK"),
        () ->
            assertTrue(
                normalizedQuickStart.contains("resolves source-backed authored inputs"),
                "quick start must describe doctor scope accurately"),
        () ->
            assertTrue(
                normalizedQuickStart.contains("preflights existing workbook-source access"),
                "quick start must mention doctor workbook-source preflight"),
        () ->
            assertTrue(
                normalizedQuickReference.contains("resolves source-backed inputs"),
                "quick reference must summarize doctor scope accurately"),
        () ->
            assertTrue(
                normalizedQuickReference.contains("preflights existing workbook-source access"),
                "quick reference must mention doctor workbook-source preflight"),
        () ->
            assertTrue(
                normalizedOperations.contains("source-backed input resolution"),
                "operations index must summarize doctor scope accurately"),
        () ->
            assertTrue(
                normalizedOperations.contains("existing workbook-source accessibility"),
                "operations index must mention doctor workbook-source preflight"));
  }

  @Test
  void requestDocsDescribeRequestOwnedPathResolutionTruthfully() throws IOException {
    String requestReference = readDoc("docs/REQUEST_AND_EXECUTION_REFERENCE.md");
    String quickReference = readDoc("docs/QUICK_REFERENCE.md");
    String quickStart = readDoc("docs/QUICK_START.md");
    String normalizedRequestReference = requestReference.replaceAll("\\s+", " ");
    String normalizedQuickReference = quickReference.replaceAll("\\s+", " ");
    String normalizedQuickStart = quickStart.replaceAll("\\s+", " ");

    assertAll(
        () ->
            assertTrue(
                normalizedRequestReference.contains(
                    "relative request-owned paths inside the JSON follow the request file directory"),
                "request reference must explain request-file-rooted relative paths"),
        () ->
            assertTrue(
                normalizedRequestReference.contains(
                    "`--request` and `--response` still resolve from the shell working directory"),
                "request reference must distinguish CLI flag path resolution"),
        () ->
            assertTrue(
                normalizedQuickReference.contains(
                    "relative request-owned paths such as `source.path`, `persistence.path`, `UTF8_FILE` / `FILE`, external workbook bindings, and signing"),
                "quick reference must summarize the request-owned path rule"),
        () ->
            assertTrue(
                normalizedQuickStart.contains(
                    "relative paths inside that JSON request follow the request file's directory"),
                "quick start must teach the request-owned path rule"));
  }

  @Test
  void errorReferenceDescribesLiveCellNotFoundPathTruthfully() throws IOException {
    String errors = readDoc("docs/ERRORS.md");

    assertAll(
        () ->
            assertTrue(
                errors.contains("execution.calculation.strategy=EVALUATE_TARGETS"),
                "error reference must describe the current CELL_NOT_FOUND path"),
        () ->
            assertFalse(
                errors.contains("Reserved. No current step raises this code"),
                "error reference must not claim CELL_NOT_FOUND is unused"));
  }

  private static String readDoc(String relativePath) throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    return Files.readString(repositoryRoot.resolve(relativePath));
  }

  private static List<CapabilityRow> capabilityRows(String markdown) {
    return markdown
        .lines()
        .filter(line -> line.startsWith("|"))
        .map(GridGrindSpreadsheetDocsAuditTest::capabilityRowOrNull)
        .filter(row -> row != null)
        .toList();
  }

  private static CapabilityRow capabilityRowOrNull(String line) {
    String[] parts = line.split("\\|", -1);
    if (parts.length != 7) {
      return null;
    }
    String capability = parts[1].trim();
    String status = parts[3].trim();
    String evidence = parts[4].trim();
    if (capability.isEmpty()
        || "Capability".equals(capability)
        || capability.startsWith(":")
        || !status.startsWith("`")) {
      return null;
    }
    return new CapabilityRow(capability, status, evidence);
  }

  private record CapabilityRow(String capability, String status, String evidence) {
    private boolean shipped() {
      return "`COMPLETE`".equals(status) || "`PARTIAL`".equals(status);
    }
  }
}
