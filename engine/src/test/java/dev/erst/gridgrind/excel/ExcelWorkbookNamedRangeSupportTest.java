package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers raw observed defined names that Apache POI does not project as live names. */
class ExcelWorkbookNamedRangeSupportTest {
  @Test
  void preservesObservedEmptyFormula() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Data");
      var name = workbook.xssfWorkbook().createName();
      name.setNameName("EmptyReference");
      var observed = workbook.names().namedRanges().getFirst();
      assertEquals("", observed.refersToFormula());
    }
  }

  @Test
  void preservesObservedInvalidSheetLabelInLocalScope() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Data");
      var name = workbook.xssfWorkbook().createName();
      name.setNameName("LocalReference");
      name.setSheetIndex(0);
      name.setRefersToFormula("Data!$A$1");
      workbook.xssfWorkbook().getCTWorkbook().getSheets().getSheetArray(0).setName("stored[]");
      var observed = workbook.names().namedRanges().getFirst();
      assertEquals(
          new NamedRangeScopeReport.Sheet(
              new dev.erst.gridgrind.core.vocabulary.SheetLabel("stored[]")),
          observed.scope());
    }
  }

  @Test
  void projectsUnmodeledObservedNamesWithoutApplyingAuthoringValidation() throws IOException {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-observed-names-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Data");
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("ModeledName"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"))));
      workbook
          .persistence()
          .save(
              workbookPath,
              WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      var definedNames = workbook.xssfWorkbook().getCTWorkbook().getDefinedNames();
      addRawName(definedNames, " ", "Data!$A$1", null, false);
      addRawName(definedNames, "HiddenRaw", "Data!$A$1", null, true);
      addRawName(definedNames, "9 legacy name", "Data!$A$1", null, false);
      addRawName(definedNames, "RawLocal", "Data!$A$1", 0L, false);
      addRawName(definedNames, "RawMissingLocal", "Data!$A$1", 1L, false);
      addRawName(definedNames, "RawLargeLocal", "Data!$A$1", 2_147_483_648L, false);

      List<NamedRangeReport> observed =
          ExcelWorkbookNamedRangeSupport.unmodeledObservedNamedRanges(workbook);

      assertEquals(
          List.of("9 legacy name", "RawLocal", "RawMissingLocal", "RawLargeLocal"),
          observed.stream().map(namedRange -> namedRange.name().value()).toList());
      assertEquals(
          new NamedRangeScopeReport.Sheet(
              new dev.erst.gridgrind.core.vocabulary.SheetLabel("Data")),
          observed.get(1).scope());
      assertTrue(observed.get(2).scope() instanceof NamedRangeScopeReport.Workbook);
      assertTrue(observed.get(3).scope() instanceof NamedRangeScopeReport.Workbook);
    }
  }

  private static void addRawName(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedNames definedNames,
      String name,
      String formula,
      Long localSheetId,
      boolean hidden) {
    var definedName = definedNames.addNewDefinedName();
    definedName.setName(name);
    definedName.setStringValue(formula);
    if (localSheetId != null) {
      definedName.setLocalSheetId(localSheetId);
    }
    if (hidden) {
      definedName.setHidden(true);
    }
  }
}
