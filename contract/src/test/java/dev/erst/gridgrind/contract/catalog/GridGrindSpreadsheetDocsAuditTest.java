package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
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
  void limitationsRegistryStatesEnforcementAndFormatScopeTruthfully() throws IOException {
    String limitations = readDoc("docs/LIMITATIONS.md");

    assertAll(
        () ->
            assertFalse(
                limitations.contains("GridGrind does not currently enforce these at request time"),
                "limitations doc must not claim blanket non-enforcement for Excel ceilings"),
        () ->
            assertTrue(
                limitations.contains(
                    "Some, such as addressed row/column bounds, are already enforced"),
                "limitations doc must explain that some Excel ceilings are preflighted"),
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
                "unsupported-features table must describe .xlsm as an out-of-scope product choice"));
  }

  private static String readDoc(String relativePath) throws IOException {
    Path repositoryRoot = RepositoryRootTestSupport.repositoryRoot();
    return Files.readString(repositoryRoot.resolve(relativePath));
  }
}
