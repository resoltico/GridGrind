package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.UUID;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

/** Integration tests for ExcelWorkbook creation, loading, and sheet access. */
class ExcelWorkbookTest {
  @Test
  void snapshotsAndPreviewExposeFormulaResults() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-engine-", ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor commandExecutor = new WorkbookCommandExecutor();
      commandExecutor.apply(
          workbook,
          List.of(
              new WorkbookCommand.CreateSheet("Budget"),
              new WorkbookCommand.AppendRow(
                  "Budget", List.of(ExcelCellValue.text("Item"), ExcelCellValue.text("Amount"))),
              new WorkbookCommand.AppendRow(
                  "Budget", List.of(ExcelCellValue.text("Hosting"), ExcelCellValue.number(49.0))),
              new WorkbookCommand.AppendRow(
                  "Budget", List.of(ExcelCellValue.text("Domain"), ExcelCellValue.number(12.0))),
              new WorkbookCommand.SetCell("Budget", "A4", ExcelCellValue.text("Total")),
              new WorkbookCommand.SetCell("Budget", "B4", ExcelCellValue.formula("SUM(B2:B3)")),
              new WorkbookCommand.EvaluateAllFormulas()));
      workbook.save(workbookPath);
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet sheet = workbook.sheet("Budget");

      ExcelCellSnapshot.FormulaSnapshot totalSnapshot =
          (ExcelCellSnapshot.FormulaSnapshot) sheet.snapshotCell("B4");
      assertEquals("FORMULA", totalSnapshot.declaredType());
      assertEquals("FORMULA", totalSnapshot.effectiveType());
      assertEquals("SUM(B2:B3)", totalSnapshot.formula());
      assertEquals(
          61.0, ((ExcelCellSnapshot.NumberSnapshot) totalSnapshot.evaluation()).numberValue());

      List<ExcelPreviewRow> preview = sheet.preview(4, 2);
      assertEquals(4, preview.size());
      assertEquals("A1", preview.get(0).cells().get(0).address());
      assertEquals(
          "Hosting",
          ((ExcelCellSnapshot.TextSnapshot) preview.get(1).cells().get(0)).stringValue());
      assertEquals("61", preview.get(3).cells().get(1).displayValue());
    }
  }

  @Test
  void managesWorkbookLifecycleAndValidation() throws IOException {
    Path workbookPath =
        Files.createTempDirectory("gridgrind-workbook-").resolve("nested/book.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Hello"));
      workbook.getOrCreateSheet("Budget").setCell("B1", ExcelCellValue.number(12.0));
      workbook.forceFormulaRecalculationOnOpen();
      workbook.save(workbookPath);

      assertEquals(1, workbook.sheetCount());
      assertEquals(List.of("Budget"), workbook.sheetNames());
      assertTrue(workbook.forceFormulaRecalculationOnOpenEnabled());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertEquals("Hello", workbook.sheet("Budget").text("A1"));
      assertEquals(12.0, workbook.sheet("Budget").number("B1"));
    }
  }

  @Test
  void createAndOpenWithFormulaEnvironmentExposeRuntimeContextAndValidateFormulaTargets()
      throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-formula-workbook-");
    Path referencedWorkbookPath = directory.resolve("rates.xlsx");
    Path workbookPath = directory.resolve("budget.xlsx");

    try (XSSFWorkbook referencedWorkbook = new XSSFWorkbook()) {
      referencedWorkbook.createSheet("Rates").createRow(0).createCell(0).setCellValue(7.5d);
      try (var outputStream = Files.newOutputStream(referencedWorkbookPath)) {
        referencedWorkbook.write(outputStream);
      }
    }

    ExcelFormulaEnvironment environment =
        new ExcelFormulaEnvironment(
            List.of(new ExcelFormulaExternalWorkbookBinding("rates.xlsx", referencedWorkbookPath)),
            ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE,
            List.of(
                new ExcelFormulaUdfToolpack(
                    "math", List.of(new ExcelFormulaUdfFunction("DOUBLE", 1, 1, "ARG1*2")))));

    try (ExcelWorkbook workbook = ExcelWorkbook.create(environment)) {
      assertTrue(workbook.formulaRuntimeContext().hasExternalWorkbookBinding("RATES.XLSX"));
      assertTrue(workbook.formulaRuntimeContext().hasUserDefinedFunction("double"));
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));
      workbook.evaluateAllFormulas();
      workbook.save(workbookPath);

      assertThrows(
          IllegalArgumentException.class,
          () -> workbook.evaluateFormulaCells(List.of(new ExcelFormulaCellTarget("Budget", "A1"))));
      assertThrows(
          InvalidCellAddressException.class,
          () -> workbook.evaluateFormulaCells(List.of(new ExcelFormulaCellTarget("Budget", ":"))));
      assertThrows(
          CellNotFoundException.class,
          () ->
              workbook.evaluateFormulaCells(List.of(new ExcelFormulaCellTarget("Budget", "Z99"))));
    }

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath, environment)) {
      assertTrue(reopened.formulaRuntimeContext().hasExternalWorkbookBinding("rates.xlsx"));
      assertTrue(reopened.formulaRuntimeContext().hasUserDefinedFunction("DOUBLE"));
    }
  }

  @Test
  void renamesDeletesAndMovesSheetsAcrossSaves() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-sheet-management-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Live"));
      workbook.getOrCreateSheet("Archive").setCell("A1", ExcelCellValue.text("Old"));
      workbook.getOrCreateSheet("Scratch");

      workbook.renameSheet("Archive", "History");
      workbook.moveSheet("History", 0);
      workbook.deleteSheet("Scratch");
      workbook.save(workbookPath);

      assertEquals(List.of("History", "Budget"), workbook.sheetNames());
      assertEquals("Old", workbook.sheet("History").text("A1"));
      assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Archive"));
    }

    assertEquals(List.of("History", "Budget"), XlsxRoundTrip.sheetOrder(workbookPath));
  }

  @Test
  void workbookSummaryUsesExplicitEmptyAndWithSheetsStates() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookReadResult.WorkbookSummary emptySummary = workbook.workbookSummary();
      WorkbookReadResult.WorkbookSummary.Empty empty =
          assertInstanceOf(WorkbookReadResult.WorkbookSummary.Empty.class, emptySummary);
      assertEquals(0, empty.sheetCount());
      assertEquals(List.of(), empty.sheetNames());

      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.getOrCreateSheet("Gamma");
      workbook.setActiveSheet("Beta");
      workbook.setSelectedSheets(List.of("Gamma", "Alpha"));
      workbook.setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN);

      WorkbookReadResult.WorkbookSummary.WithSheets summary =
          assertInstanceOf(
              WorkbookReadResult.WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals(List.of("Alpha", "Beta", "Gamma"), summary.sheetNames());
      assertEquals("Gamma", summary.activeSheetName());
      assertEquals(List.of("Alpha", "Gamma"), summary.selectedSheetNames());
    }
  }

  @Test
  void sheetStateRoundTripsAcrossSaveAndReopen() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-sheet-state-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha").setCell("A1", ExcelCellValue.text("Live"));
      workbook.getOrCreateSheet("Beta");
      workbook.getOrCreateSheet("Gamma");
      workbook.copySheet("Alpha", "Replica", new ExcelSheetCopyPosition.AtIndex(1));
      workbook.setActiveSheet("Beta");
      workbook.setSelectedSheets(List.of("Gamma", "Alpha"));
      workbook.setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN);
      workbook.setSheetVisibility("Replica", ExcelSheetVisibility.VERY_HIDDEN);
      workbook.setSheetProtection("Alpha", protectionSettings());
      workbook.save(workbookPath);
    }

    assertEquals(
        List.of("Alpha", "Replica", "Beta", "Gamma"), XlsxRoundTrip.sheetOrder(workbookPath));
    assertEquals("Gamma", XlsxRoundTrip.activeSheetName(workbookPath));
    assertEquals(List.of("Alpha", "Gamma"), XlsxRoundTrip.selectedSheetNames(workbookPath));
    assertEquals(ExcelSheetVisibility.HIDDEN, XlsxRoundTrip.sheetVisibility(workbookPath, "Beta"));
    assertEquals(
        ExcelSheetVisibility.VERY_HIDDEN, XlsxRoundTrip.sheetVisibility(workbookPath, "Replica"));
    assertEquals(
        new WorkbookReadResult.SheetProtection.Protected(protectionSettings()),
        XlsxRoundTrip.sheetProtection(workbookPath, "Alpha"));

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadResult.WorkbookSummary.WithSheets summary =
          assertInstanceOf(
              WorkbookReadResult.WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals("Gamma", summary.activeSheetName());
      assertEquals(List.of("Alpha", "Gamma"), summary.selectedSheetNames());

      WorkbookReadResult.SheetSummary alphaSummary = workbook.sheetSummary("Alpha");
      WorkbookReadResult.SheetSummary betaSummary = workbook.sheetSummary("Beta");
      WorkbookReadResult.SheetSummary replicaSummary = workbook.sheetSummary("Replica");
      assertEquals(ExcelSheetVisibility.VISIBLE, alphaSummary.visibility());
      assertEquals(
          new WorkbookReadResult.SheetProtection.Protected(protectionSettings()),
          alphaSummary.protection());
      assertEquals(ExcelSheetVisibility.HIDDEN, betaSummary.visibility());
      assertInstanceOf(
          WorkbookReadResult.SheetProtection.Unprotected.class, betaSummary.protection());
      assertEquals(ExcelSheetVisibility.VERY_HIDDEN, replicaSummary.visibility());
    }
  }

  @Test
  void clearSheetProtectionIsIdempotentForUnprotectedSheets() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-clear-sheet-protection-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.clearSheetProtection("Alpha");
      workbook.save(workbookPath);
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertInstanceOf(
          WorkbookReadResult.SheetProtection.Unprotected.class,
          workbook.sheetSummary("Alpha").protection());
    }
  }

  @Test
  void clearSheetProtectionRemovesExistingProtectionAcrossSaveAndReopen() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-clear-existing-sheet-protection-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.setSheetProtection("Alpha", protectionSettings());
      workbook.clearSheetProtection("Alpha");
      workbook.save(workbookPath);
    }

    assertEquals(
        new WorkbookReadResult.SheetProtection.Unprotected(),
        XlsxRoundTrip.sheetProtection(workbookPath, "Alpha"));

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertInstanceOf(
          WorkbookReadResult.SheetProtection.Unprotected.class,
          workbook.sheetSummary("Alpha").protection());
    }
  }

  @Test
  void copySheetPreservesSupportedLocalStructuresAndCopiesSheetScopedRangeNames()
      throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Source");
      workbook.sheet("Source").setCell("A1", ExcelCellValue.text("Item"));
      workbook.sheet("Source").setCell("B1", ExcelCellValue.text("Amount"));
      workbook.sheet("Source").setCell("A2", ExcelCellValue.text("Hosting"));
      workbook.sheet("Source").setCell("B2", ExcelCellValue.number(49.0));
      workbook.sheet("Source").setCell("B3", ExcelCellValue.formula("SUM(B2:B2)"));
      workbook.sheet("Source").setHyperlink("A2", new ExcelHyperlink.Url("https://example.com/h"));
      workbook.sheet("Source").setComment("A2", new ExcelComment("Review", "GridGrind", false));
      workbook
          .sheet("Source")
          .setDataValidation(
              "C2:C4",
              new ExcelDataValidationDefinition(
                  new ExcelDataValidationRule.WholeNumber(
                      ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                  false,
                  false,
                  null,
                  null));
      workbook
          .sheet("Source")
          .setConditionalFormatting(
              new ExcelConditionalFormattingBlockDefinition(
                  List.of("B2:B4"),
                  List.of(
                      new ExcelConditionalFormattingRule.FormulaRule(
                          "B2>0",
                          true,
                          new ExcelDifferentialStyle(
                              "0.00", true, null, null, "#102030", null, null, "#E0F0AA", null)))));
      workbook.sheet("Source").mergeCells("A1:B1");
      workbook.sheet("Source").setPane(new ExcelSheetPane.Frozen(1, 1, 1, 1));
      workbook.sheet("Source").setZoom(140);
      workbook
          .sheet("Source")
          .setPrintLayout(
              new ExcelPrintLayout(
                  new ExcelPrintLayout.Area.Range("A1:C20"),
                  ExcelPrintOrientation.LANDSCAPE,
                  new ExcelPrintLayout.Scaling.Fit(1, 0),
                  new ExcelPrintLayout.TitleRows.Band(0, 0),
                  new ExcelPrintLayout.TitleColumns.None(),
                  new ExcelHeaderFooterText("Source", "", ""),
                  new ExcelHeaderFooterText("", "&P", "")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "LocalBudget",
              new ExcelNamedRangeScope.SheetScope("Source"),
              new ExcelNamedRangeTarget("Source", "A1:B3")));
      workbook.copySheet("Source", "Replica", new ExcelSheetCopyPosition.AppendAtEnd());
      workbook.save(workbookPath);
    }

    assertEquals(
        new ExcelSheetPane.Frozen(1, 1, 1, 1), XlsxRoundTrip.pane(workbookPath, "Replica"));
    assertEquals(140, XlsxRoundTrip.zoomPercent(workbookPath, "Replica"));
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        XlsxRoundTrip.printLayout(workbookPath, "Replica").orientation());
    assertEquals(
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "LocalBudget",
                new ExcelNamedRangeScope.SheetScope("Source"),
                "Source!$A$1:Source!$B$3",
                new ExcelNamedRangeTarget("Source", "A1:B3")),
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "LocalBudget",
                new ExcelNamedRangeScope.SheetScope("Replica"),
                "Replica!$A$1:Replica!$B$3",
                new ExcelNamedRangeTarget("Replica", "A1:B3"))),
        XlsxRoundTrip.namedRanges(workbookPath));
    assertEquals(
        List.of(
            new ExcelDataValidationSnapshot.Supported(
                List.of("C2:C4"),
                new ExcelDataValidationDefinition(
                    new ExcelDataValidationRule.WholeNumber(
                        ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                    false,
                    true,
                    null,
                    null))),
        XlsxRoundTrip.dataValidations(workbookPath, "Replica"));

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertEquals("Item", workbook.sheet("Replica").text("A1"));
      assertEquals(
          "https://example.com/h",
          workbook
              .sheet("Replica")
              .snapshotCell("A2")
              .metadata()
              .hyperlink()
              .orElseThrow()
              .target());
      assertEquals(
          "Review",
          workbook.sheet("Replica").snapshotCell("A2").metadata().comment().orElseThrow().text());
      assertEquals(
          List.of("A1:B1"),
          workbook.sheet("Replica").mergedRegions().stream()
              .map(WorkbookReadResult.MergedRegion::range)
              .toList());
      assertEquals(
          List.of("B2:B4"),
          workbook
              .sheet("Replica")
              .conditionalFormatting(new ExcelRangeSelection.All())
              .getFirst()
              .ranges());
    }
  }

  @Test
  void copySheetRejectsUnsupportedSourceStructuresAndVisibilityRulesStayHonest()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN);
      assertDoesNotThrow(() -> workbook.setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN));

      IllegalArgumentException lastVisible =
          assertThrows(
              IllegalArgumentException.class,
              () -> workbook.setSheetVisibility("Alpha", ExcelSheetVisibility.HIDDEN));
      assertEquals("cannot hide the last visible sheet 'Alpha'", lastVisible.getMessage());

      IllegalArgumentException deleteLastVisible =
          assertThrows(IllegalArgumentException.class, () -> workbook.deleteSheet("Alpha"));
      assertEquals("cannot delete the last visible sheet 'Alpha'", deleteLastVisible.getMessage());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Tables");
      workbook.sheet("Tables").setCell("A1", ExcelCellValue.text("Name"));
      workbook.sheet("Tables").setCell("B1", ExcelCellValue.text("Value"));
      workbook.sheet("Tables").setCell("A2", ExcelCellValue.text("Ops"));
      workbook.sheet("Tables").setCell("B2", ExcelCellValue.number(1.0));
      workbook.setTable(
          new ExcelTableDefinition(
              "OpsTable", "Tables", "A1:B2", false, new ExcelTableStyle.None()));

      workbook.copySheet("Tables", "Tables Copy", new ExcelSheetCopyPosition.AppendAtEnd());

      assertEquals(List.of("Tables", "Tables Copy"), workbook.sheetNames());
      assertEquals(1, workbook.xssfWorkbook().getSheet("Tables").getTables().size());
      assertEquals(1, workbook.xssfWorkbook().getSheet("Tables Copy").getTables().size());
      assertEquals(
          "OpsTable", workbook.xssfWorkbook().getSheet("Tables").getTables().getFirst().getName());
      assertEquals(
          "OpsTable_Copy2",
          workbook.xssfWorkbook().getSheet("Tables Copy").getTables().getFirst().getName());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("FormulaNames");
      Name localFormula = workbook.xssfWorkbook().createName();
      localFormula.setNameName("LocalFormula");
      localFormula.setSheetIndex(workbook.xssfWorkbook().getSheetIndex("FormulaNames"));
      localFormula.setRefersToFormula("SUM(FormulaNames!$A$1:$A$2)");

      workbook.copySheet(
          "FormulaNames", "FormulaNames Copy", new ExcelSheetCopyPosition.AppendAtEnd());

      assertEquals(
          List.of(
              new ExcelNamedRangeSnapshot.FormulaSnapshot(
                  "LocalFormula",
                  new ExcelNamedRangeScope.SheetScope("FormulaNames"),
                  "SUM(FormulaNames!$A$1:$A$2)"),
              new ExcelNamedRangeSnapshot.FormulaSnapshot(
                  "LocalFormula",
                  new ExcelNamedRangeScope.SheetScope("FormulaNames Copy"),
                  "SUM('FormulaNames Copy'!$A$1:$A$2)")),
          workbook.namedRanges());
    }
  }

  @Test
  void persistsStructuralLayoutOperationsAcrossSaves() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-layout-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.setCell("A1", ExcelCellValue.text("Quarterly"));
      budget.mergeCells("A1:B1");
      budget.setColumnWidth(0, 1, 16.0);
      budget.setRowHeight(0, 0, 28.5);
      budget.setPane(new ExcelSheetPane.Frozen(1, 1, 1, 1));
      budget.setZoom(125);
      budget.setPrintLayout(
          new ExcelPrintLayout(
              new ExcelPrintLayout.Area.Range("A1:B12"),
              ExcelPrintOrientation.LANDSCAPE,
              new ExcelPrintLayout.Scaling.Fit(1, 0),
              new ExcelPrintLayout.TitleRows.Band(0, 0),
              new ExcelPrintLayout.TitleColumns.Band(0, 0),
              new ExcelHeaderFooterText("Budget", "", ""),
              new ExcelHeaderFooterText("", "Page &P", "")));
      workbook.save(workbookPath);
    }

    assertEquals(List.of("A1:B1"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), XlsxRoundTrip.pane(workbookPath, "Budget"));
    assertEquals(125, XlsxRoundTrip.zoomPercent(workbookPath, "Budget"));
    assertEquals(
        new ExcelPrintLayout.Area.Range("A1:B12"),
        XlsxRoundTrip.printLayout(workbookPath, "Budget").printArea());
  }

  @Test
  void wrapsWorkbookWideFormulaEvaluationFailures() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook =
            new ExcelWorkbook(
                poiWorkbook,
                FormulaRuntimeTestDouble.failingEvaluation(
                    poiWorkbook.getCreationHelper().createFormulaEvaluator(),
                    new org.apache.poi.ss.formula.FakeFormulaFailure("bad formula")))) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.formula("1+1"));

      InvalidFormulaException exception =
          assertThrows(InvalidFormulaException.class, workbook::evaluateAllFormulas);
      assertEquals("Budget", exception.sheetName());
      assertEquals("A1", exception.address());
      assertEquals("1+1", exception.formula());
    }
  }

  @Test
  void adaptsDirectPoiFormulaEvaluatorsForWorkbookWideEvaluation() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook =
            new ExcelWorkbook(
                poiWorkbook, poiWorkbook.getCreationHelper().createFormulaEvaluator())) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.formula("1+1"));

      workbook.evaluateAllFormulas();

      ExcelCellSnapshot.FormulaSnapshot snapshot =
          (ExcelCellSnapshot.FormulaSnapshot) workbook.sheet("Budget").snapshotCell("A1");
      assertEquals("2", snapshot.displayValue());
    }
  }

  @Test
  void validatesWorkbookInputsAndMissingResources() throws IOException {
    assertThrows(NullPointerException.class, () -> ExcelWorkbook.open(null));

    Path missingPath = Files.createTempDirectory("gridgrind-missing-").resolve("missing.xlsx");
    WorkbookNotFoundException missingWorkbook =
        assertThrows(WorkbookNotFoundException.class, () -> ExcelWorkbook.open(missingPath));
    assertTrue(missingWorkbook.getMessage().contains("Workbook does not exist"));
    assertEquals(missingPath.toAbsolutePath(), missingWorkbook.workbookPath());

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      assertThrows(NullPointerException.class, () -> workbook.getOrCreateSheet(null));
      assertThrows(IllegalArgumentException.class, () -> workbook.getOrCreateSheet(" "));
      assertThrows(NullPointerException.class, () -> workbook.sheet(null));
      assertThrows(IllegalArgumentException.class, () -> workbook.sheet(" "));
      SheetNotFoundException missingSheet =
          assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Missing"));
      assertEquals("Missing", missingSheet.sheetName());
      workbook.getOrCreateSheet("Budget");
      workbook.getOrCreateSheet("Archive");
      assertSame(workbook, workbook.renameSheet("Budget", "Budget"));
      assertThrows(NullPointerException.class, () -> workbook.renameSheet(null, "Summary"));
      assertThrows(IllegalArgumentException.class, () -> workbook.renameSheet(" ", "Summary"));
      assertThrows(NullPointerException.class, () -> workbook.renameSheet("Budget", null));
      assertThrows(IllegalArgumentException.class, () -> workbook.renameSheet("Budget", " "));
      assertThrows(SheetNotFoundException.class, () -> workbook.renameSheet("Missing", "Summary"));
      assertThrows(IllegalArgumentException.class, () -> workbook.renameSheet("Budget", "Archive"));
      assertThrows(
          IllegalArgumentException.class, () -> workbook.renameSheet("Budget", "Bad/Name"));
      assertThrows(NullPointerException.class, () -> workbook.deleteSheet(null));
      assertThrows(IllegalArgumentException.class, () -> workbook.deleteSheet(" "));
      assertThrows(SheetNotFoundException.class, () -> workbook.deleteSheet("Missing"));
      workbook.deleteSheet("Archive"); // leaves only Budget; next delete must be rejected
      IllegalArgumentException lastSheet =
          assertThrows(IllegalArgumentException.class, () -> workbook.deleteSheet("Budget"));
      assertTrue(lastSheet.getMessage().contains("at least one sheet"));
      workbook.getOrCreateSheet("Archive"); // restore two-sheet state for moveSheet tests
      assertThrows(NullPointerException.class, () -> workbook.moveSheet(null, 0));
      assertThrows(IllegalArgumentException.class, () -> workbook.moveSheet(" ", 0));
      assertThrows(SheetNotFoundException.class, () -> workbook.moveSheet("Missing", 0));
      IllegalArgumentException negativeIndex =
          assertThrows(IllegalArgumentException.class, () -> workbook.moveSheet("Budget", -1));
      assertTrue(negativeIndex.getMessage().contains("workbook has"));
      IllegalArgumentException tooLargeIndex =
          assertThrows(IllegalArgumentException.class, () -> workbook.moveSheet("Budget", 2));
      assertTrue(tooLargeIndex.getMessage().contains("valid positions are 0 to 1"));
      assertThrows(NullPointerException.class, () -> workbook.save(null));
    }
  }

  @Test
  void rejectsNonXlsxWorkbookFiles() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-legacy-", ".xls");

    try (HSSFWorkbook workbook = new HSSFWorkbook();
        var outputStream = Files.newOutputStream(workbookPath)) {
      workbook.createSheet("Budget").createRow(0).createCell(0).setCellValue("Legacy");
      workbook.write(outputStream);
    }

    IllegalArgumentException exception =
        assertThrows(IllegalArgumentException.class, () -> ExcelWorkbook.open(workbookPath));
    assertEquals("Only .xlsx workbooks are supported", exception.getMessage());
  }

  @SuppressWarnings({"PMD.CloseResource", "PMD.UseTryWithResources"})
  @Test
  void validatesFormulaEnvironmentOverloadsAndCloseFailureAggregation() throws Exception {
    ExcelFormulaEnvironment defaults = ExcelFormulaEnvironment.defaults();

    assertThrows(NullPointerException.class, () -> ExcelWorkbook.open(null, defaults));

    Path missingPath =
        Files.createTempDirectory("gridgrind-missing-env-").resolve("missing-with-env.xlsx");
    WorkbookNotFoundException missingWorkbook =
        assertThrows(
            WorkbookNotFoundException.class, () -> ExcelWorkbook.open(missingPath, defaults));
    assertEquals(missingPath.toAbsolutePath(), missingWorkbook.workbookPath());

    Path legacyWorkbookPath = Files.createTempFile("gridgrind-legacy-env-", ".xls");
    try (HSSFWorkbook workbook = new HSSFWorkbook();
        var outputStream = Files.newOutputStream(legacyWorkbookPath)) {
      workbook.createSheet("Budget").createRow(0).createCell(0).setCellValue("Legacy");
      workbook.write(outputStream);
    }
    IllegalArgumentException unsupportedFormat =
        assertThrows(
            IllegalArgumentException.class, () -> ExcelWorkbook.open(legacyWorkbookPath, defaults));
    assertEquals("Only .xlsx workbooks are supported", unsupportedFormat.getMessage());

    try (ExcelWorkbook workbook = ExcelWorkbook.create(null)) {
      assertEquals(
          ExcelFormulaEnvironment.defaults().runtimeContext(), workbook.formulaRuntimeContext());
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Header"));
      assertThrows(
          CellNotFoundException.class,
          () -> workbook.evaluateFormulaCells(List.of(new ExcelFormulaCellTarget("Budget", "B1"))));
    }

    ThrowingCloseWorkbook workbookDelegate = new ThrowingCloseWorkbook("workbook close failure");
    ExcelWorkbook workbook =
        new ExcelWorkbook(
            workbookDelegate, new ThrowingCloseFormulaRuntime("runtime close failure"));
    try {
      IOException failure = assertThrows(IOException.class, workbook::close);
      assertEquals("runtime close failure", failure.getMessage());
      assertEquals(1, failure.getSuppressed().length);
      assertEquals("workbook close failure", failure.getSuppressed()[0].getMessage());
    } finally {
      workbookDelegate.disableCloseFailure();
      workbookDelegate.close();
    }

    ThrowingCloseWorkbook workbookOnlyFailure = new ThrowingCloseWorkbook("workbook only failure");
    ExcelWorkbook workbookWithHealthyRuntime =
        new ExcelWorkbook(
            workbookOnlyFailure, workbookOnlyFailure.getCreationHelper().createFormulaEvaluator());
    try {
      IOException failure = assertThrows(IOException.class, workbookWithHealthyRuntime::close);
      assertEquals("workbook only failure", failure.getMessage());
      assertEquals(0, failure.getSuppressed().length);
    } finally {
      workbookOnlyFailure.disableCloseFailure();
      workbookOnlyFailure.close();
    }
  }

  @Test
  void clearFormulaCachesRemovesInlineStringAndTypeMetadata() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.formula("\"hello\""));
      workbook.evaluateAllFormulas();

      org.apache.poi.xssf.usermodel.XSSFCell cell =
          workbook.xssfWorkbook().getSheet("Budget").getRow(0).getCell(0);
      var ctCell = cell.getCTCell();
      if (ctCell.isSetV()) {
        ctCell.unsetV();
      }
      ctCell.addNewIs().setT("hello");
      ctCell.setT(org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType.INLINE_STR);
      assertTrue(ctCell.isSetIs());
      assertTrue(ctCell.isSetT());

      workbook.clearFormulaCaches();

      assertFalse(ctCell.isSetV());
      assertFalse(ctCell.isSetIs());
      assertFalse(ctCell.isSetT());
    }
  }

  @Test
  void clearFormulaCachesLeavesUnsetTypeMetadataUnset() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.formula("1+1"));
      workbook.evaluateAllFormulas();

      org.apache.poi.xssf.usermodel.XSSFCell cell =
          workbook.xssfWorkbook().getSheet("Budget").getRow(0).getCell(0);
      var ctCell = cell.getCTCell();
      if (ctCell.isSetT()) {
        ctCell.unsetT();
      }
      assertFalse(ctCell.isSetT());

      workbook.clearFormulaCaches();

      assertFalse(ctCell.isSetT());
    }
  }

  @Test
  void savesToPathsWithoutParentDirectories() throws IOException {
    Path relativePath = Path.of("gridgrind-relative-" + UUID.randomUUID() + ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Hello"));
      workbook.save(relativePath);
      assertTrue(Files.exists(relativePath));
    } finally {
      Files.deleteIfExists(relativePath);
    }
  }

  @Test
  void roundTripHelpersInspectSavedWorkbookStructure() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-structure-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var budget = workbook.createSheet("Budget");
      workbook.createSheet("Summary");

      budget.addMergedRegion(CellRangeAddress.valueOf("A1:B2"));
      budget.setColumnWidth(0, 4096);
      budget.createRow(0).setHeightInPoints(28.5f);
      budget.createFreezePane(1, 2, 3, 4);

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    assertEquals(List.of("Budget", "Summary"), XlsxRoundTrip.sheetOrder(workbookPath));
    assertEquals(List.of("A1:B2"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(new ExcelSheetPane.Frozen(1, 2, 3, 4), XlsxRoundTrip.pane(workbookPath, "Budget"));
  }

  @Test
  void savesAndReopensFormattingDepthStyles() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-style-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Item"));
      workbook
          .sheet("Budget")
          .applyStyle(
              "A1",
              new ExcelCellStyle(
                  null,
                  new ExcelCellAlignment(
                      true,
                      ExcelHorizontalAlignment.CENTER,
                      ExcelVerticalAlignment.TOP,
                      null,
                      null),
                  new ExcelCellFont(
                      true,
                      false,
                      "Aptos",
                      ExcelFontHeight.fromPoints(new BigDecimal("11.5")),
                      new ExcelColor("#1F4E78"),
                      true,
                      true),
                  new ExcelCellFill(ExcelFillPattern.SOLID, "#FFF2CC", null),
                  new ExcelBorder(
                      new ExcelBorderSide(ExcelBorderStyle.THIN),
                      null,
                      new ExcelBorderSide(ExcelBorderStyle.DOUBLE),
                      null,
                      null),
                  null));
      workbook.save(workbookPath);
    }

    ExcelCellStyleSnapshot style = XlsxRoundTrip.cellStyle(workbookPath, "Budget", "A1");
    assertTrue(style.font().bold());
    assertFalse(style.font().italic());
    assertTrue(style.alignment().wrapText());
    assertEquals(ExcelHorizontalAlignment.CENTER, style.alignment().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, style.alignment().verticalAlignment());
    assertEquals("Aptos", style.font().fontName());
    assertEquals(new BigDecimal("11.5"), style.font().fontHeight().points());
    assertEquals(rgb("#1F4E78"), style.font().fontColor());
    assertTrue(style.font().underline());
    assertTrue(style.font().strikeout());
    assertEquals(rgb("#FFF2CC"), style.fill().foregroundColor());
    assertEquals(ExcelBorderStyle.THIN, style.border().top().style());
    assertEquals(ExcelBorderStyle.DOUBLE, style.border().right().style());
    assertEquals(ExcelBorderStyle.THIN, style.border().bottom().style());
    assertEquals(ExcelBorderStyle.THIN, style.border().left().style());
  }

  @Test
  void persistsHyperlinksCommentsAndNamedRangesAcrossSaves() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-authoring-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.setCell("A1", ExcelCellValue.text("Report"));
      budget.setCell("B4", ExcelCellValue.number(61.0));
      budget.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      budget.setComment("A1", new ExcelComment("Review", "GridGrind", true));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "B4")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "LocalItem",
              new ExcelNamedRangeScope.SheetScope("Budget"),
              new ExcelNamedRangeTarget("Budget", "A1:B2")));
      workbook.save(workbookPath);
    }

    ExcelCellMetadataSnapshot metadata = XlsxRoundTrip.cellMetadata(workbookPath, "Budget", "A1");
    assertEquals(
        new ExcelHyperlink.Url("https://example.com/report"), metadata.hyperlink().orElseThrow());
    assertEquals(
        new ExcelComment("Review", "GridGrind", true),
        metadata.comment().orElseThrow().toPlainComment());
    List<ExcelNamedRangeSnapshot> namedRanges = XlsxRoundTrip.namedRanges(workbookPath);
    assertEquals(2, namedRanges.size());
    assertTrue(
        namedRanges.contains(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                namedRanges.stream()
                    .filter(namedRange -> "BudgetTotal".equals(namedRange.name()))
                    .findFirst()
                    .orElseThrow()
                    .refersToFormula(),
                new ExcelNamedRangeTarget("Budget", "B4"))));
    assertTrue(
        namedRanges.contains(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "LocalItem",
                new ExcelNamedRangeScope.SheetScope("Budget"),
                namedRanges.stream()
                    .filter(namedRange -> "LocalItem".equals(namedRange.name()))
                    .findFirst()
                    .orElseThrow()
                    .refersToFormula(),
                new ExcelNamedRangeTarget("Budget", "A1:B2"))));

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertEquals(2, workbook.namedRangeCount());
      assertEquals(2, workbook.namedRanges().size());
      workbook.deleteNamedRange("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
      assertEquals(1, workbook.namedRangeCount());
    }
  }

  @Test
  void persistsLatestHyperlinkTargetAfterRepeatedWrites() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-hyperlink-replace-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("C");
      sheet.setHyperlink("F18", new ExcelHyperlink.Email("Report_Value@example.com"));
      sheet.setHyperlink("F18", new ExcelHyperlink.Email("Summary.Total@example.com"));
      workbook.save(workbookPath);
    }

    ExcelCellMetadataSnapshot metadata = XlsxRoundTrip.cellMetadata(workbookPath, "C", "F18");
    assertEquals(
        new ExcelHyperlink.Email("Summary.Total@example.com"), metadata.hyperlink().orElseThrow());
  }

  @Test
  void replacesNamedRangesAndExposesOnlySupportedNamedRangeSnapshots() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook =
            new ExcelWorkbook(
                poiWorkbook,
                FormulaRuntimeTestDouble.delegating(
                    poiWorkbook.getCreationHelper().createFormulaEvaluator()))) {
      poiWorkbook.createSheet("Budget");

      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "B4")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "C1")));

      Name formulaName = poiWorkbook.createName();
      formulaName.setNameName("BudgetRollup");
      formulaName.setRefersToFormula("SUM(Budget!$B$2:$B$3)");
      Name internalLowerDefinedName = poiWorkbook.createName();
      internalLowerDefinedName.setNameName("_xlnm.Print_Area");
      internalLowerDefinedName.setRefersToFormula("Budget!$A$1");
      Name internalUpperDefinedName = poiWorkbook.createName();
      internalUpperDefinedName.setNameName("_XLNM.PRINT_TITLES");
      internalUpperDefinedName.setRefersToFormula("Budget!$A$1");

      Name workbookScoped = poiWorkbook.getName("BudgetTotal");
      Name sheetScoped = poiWorkbook.createName();
      sheetScoped.setNameName("LocalItem");
      sheetScoped.setSheetIndex(poiWorkbook.getSheetIndex("Budget"));
      sheetScoped.setRefersToFormula("Budget!$A$1");

      assertTrue(ExcelWorkbook.shouldExpose(workbookScoped));
      assertFalse(ExcelWorkbook.shouldExpose(null, false, false));
      assertFalse(ExcelWorkbook.shouldExpose("HiddenBudgetTotal", false, true));
      assertFalse(ExcelWorkbook.shouldExpose("_xlnm.Print_Area", false, false));
      assertFalse(ExcelWorkbook.shouldExpose("_XLNM.PRINT_TITLES", false, false));
      assertFalse(ExcelWorkbook.shouldExpose("BudgetFn", true, false));

      assertTrue(workbook.scopeMatches(workbookScoped, new ExcelNamedRangeScope.WorkbookScope()));
      assertFalse(
          workbook.scopeMatches(workbookScoped, new ExcelNamedRangeScope.SheetScope("Budget")));
      assertTrue(workbook.scopeMatches(sheetScoped, new ExcelNamedRangeScope.SheetScope("Budget")));
      assertFalse(workbook.scopeMatches(sheetScoped, new ExcelNamedRangeScope.WorkbookScope()));
      assertThrows(
          SheetNotFoundException.class,
          () -> workbook.scopeMatches(sheetScoped, new ExcelNamedRangeScope.SheetScope("Missing")));

      assertEquals(
          List.of(
              new ExcelNamedRangeSnapshot.RangeSnapshot(
                  "BudgetTotal",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  "Budget!$C$1",
                  new ExcelNamedRangeTarget("Budget", "C1")),
              new ExcelNamedRangeSnapshot.FormulaSnapshot(
                  "BudgetRollup",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  "SUM(Budget!$B$2:$B$3)"),
              new ExcelNamedRangeSnapshot.RangeSnapshot(
                  "LocalItem",
                  new ExcelNamedRangeScope.SheetScope("Budget"),
                  "Budget!$A$1",
                  new ExcelNamedRangeTarget("Budget", "A1"))),
          workbook.namedRanges());
    }
  }

  @Test
  void validatesNamedRangeOperations() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");

      assertThrows(NullPointerException.class, () -> workbook.setNamedRange(null));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              workbook.setNamedRange(
                  new ExcelNamedRangeDefinition(
                      "A1",
                      new ExcelNamedRangeScope.WorkbookScope(),
                      new ExcelNamedRangeTarget("Budget", "B4"))));
      assertThrows(
          NullPointerException.class,
          () -> workbook.deleteNamedRange(null, new ExcelNamedRangeScope.WorkbookScope()));
      assertThrows(
          NullPointerException.class, () -> workbook.deleteNamedRange("BudgetTotal", null));
      assertThrows(
          NamedRangeNotFoundException.class,
          () -> workbook.deleteNamedRange("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope()));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              workbook.setNamedRange(
                  new ExcelNamedRangeDefinition(
                      "BudgetTotal",
                      new ExcelNamedRangeScope.SheetScope("Missing"),
                      new ExcelNamedRangeTarget("Missing", "A1"))));
    }
  }

  @Test
  void rejectsSheetNamesExceeding31Characters() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      String exactly31 = "A".repeat(31);
      String tooLong = "A".repeat(32);

      assertDoesNotThrow(() -> workbook.getOrCreateSheet(exactly31));
      assertThrows(IllegalArgumentException.class, () -> workbook.getOrCreateSheet(tooLong));
      assertThrows(IllegalArgumentException.class, () -> workbook.sheet(tooLong));
    }
  }

  @Test
  void saveCanonicalizesAmbiguousPoiColumnOutlineDefinitions() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-column-save-", ".xlsx");

    try {
      try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
        XSSFSheet sheet = workbook.xssfWorkbook().createSheet("Budget");

        sheet.groupColumn(2, 3);
        for (int repetition = 0; repetition < 6; repetition++) {
          sheet.groupColumn(2, 2);
        }
        sheet.groupColumn(1, 3);

        assertFalse(
            columnDefinitionsAreCanonical(sheet),
            "raw Apache POI grouping should leave ambiguous overlapping column definitions");

        workbook.save(workbookPath);
      }

      try (XSSFWorkbook reopenedWorkbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
        XSSFSheet reopenedSheet = reopenedWorkbook.getSheet("Budget");

        assertTrue(columnDefinitionsAreCanonical(reopenedSheet));
        assertEquals(1, reopenedSheet.getColumnOutlineLevel(1));
        assertEquals(7, reopenedSheet.getColumnOutlineLevel(2));
        assertEquals(2, reopenedSheet.getColumnOutlineLevel(3));
      }
    } finally {
      Files.deleteIfExists(workbookPath);
    }
  }

  private static ExcelSheetProtectionSettings protectionSettings() {
    return new ExcelSheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }

  private static ExcelColorSnapshot rgb(String rgb) {
    return new ExcelColorSnapshot(rgb);
  }

  private static boolean columnDefinitionsAreCanonical(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().sizeOfColsArray() != 1) {
      return false;
    }
    boolean[] seenColumns = new boolean[ExcelColumnSpan.MAX_COLUMN_INDEX + 1];
    for (CTCol col : sheet.getCTWorksheet().getColsArray(0).getColList()) {
      if (col.getMin() != col.getMax()) {
        return false;
      }
      for (int columnIndex = (int) col.getMin() - 1;
          columnIndex <= (int) col.getMax() - 1;
          columnIndex++) {
        if (seenColumns[columnIndex]) {
          return false;
        }
        seenColumns[columnIndex] = true;
      }
    }
    return true;
  }

  /** Test-only workbook that can fail on close so lifecycle aggregation is observable. */
  private static final class ThrowingCloseWorkbook extends XSSFWorkbook {
    private final String message;
    private boolean failOnClose = true;

    private ThrowingCloseWorkbook(String message) {
      this.message = message;
    }

    private void disableCloseFailure() {
      failOnClose = false;
    }

    @Override
    public void close() throws IOException {
      if (failOnClose) {
        throw new IOException(message);
      }
      super.close();
    }
  }

  /** Test-only runtime that surfaces close failures without adding evaluation behavior. */
  private record ThrowingCloseFormulaRuntime(String message) implements ExcelFormulaRuntime {
    @Override
    public org.apache.poi.ss.usermodel.CellValue evaluate(org.apache.poi.ss.usermodel.Cell cell) {
      return null;
    }

    @Override
    public org.apache.poi.ss.usermodel.CellType evaluateFormulaCell(
        org.apache.poi.ss.usermodel.Cell cell) {
      return org.apache.poi.ss.usermodel.CellType._NONE;
    }

    @Override
    public void clearCachedResults() {}

    @Override
    public String displayValue(
        org.apache.poi.ss.usermodel.DataFormatter formatter,
        org.apache.poi.ss.usermodel.Cell cell) {
      return "";
    }

    @Override
    public ExcelFormulaRuntimeContext context() {
      return ExcelFormulaEnvironment.defaults().runtimeContext();
    }

    @Override
    public void close() throws IOException {
      throw new IOException(message);
    }
  }
}
