package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookAnalyzer finding-bearing analysis reads. */
class WorkbookAnalyzerTest {
  @Test
  void executesEveryAnalysisCommandAgainstWorkbookState() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-analyzer-", ".xlsx");
    writeAnalyzerFixture(workbookPath);

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelSheet budget = workbook.sheet("Budget");
      budget.setHyperlink("A1", new ExcelHyperlink.Document("Missing!A1"));
      budget.setHyperlink("A2", new ExcelHyperlink.Document("Budget!A1:"));
      budget.setHyperlink("A3", new ExcelHyperlink.File("missing-supporting-doc.pdf"));
      budget.setCell("D1", ExcelCellValue.text("Owner"));
      budget.setCell("E1", ExcelCellValue.text("Task"));
      budget.setCell("D2", ExcelCellValue.text("Ada"));
      budget.setCell("E2", ExcelCellValue.text("Queue"));
      budget.setCell("D3", ExcelCellValue.text("Lin"));
      budget.setCell("E3", ExcelCellValue.text("Pack"));
      budget.setConditionalFormatting(
          new ExcelConditionalFormattingBlockDefinition(
              List.of("B2:B3"),
              List.of(
                  new ExcelConditionalFormattingRule.FormulaRule(
                      "B2>0",
                      false,
                      new ExcelDifferentialStyle(
                          "0.00", null, null, null, null, null, null, null, null)))));
      workbook.setTable(
          new ExcelTableDefinition("Queue", "Budget", "D1:E3", false, new ExcelTableStyle.None()));
      budget.xssfSheet().getTables().getFirst().getCTTable().getAutoFilter().setRef("D1:E2");

      WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

      WorkbookReadResult.FormulaHealthResult formulaHealth =
          cast(
              WorkbookReadResult.FormulaHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeFormulaHealth(
                      "formulaHealth", new ExcelSheetSelection.All())));
      WorkbookReadResult.HyperlinkHealthResult hyperlinkHealth =
          cast(
              WorkbookReadResult.HyperlinkHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeHyperlinkHealth(
                      "hyperlinkHealth", new ExcelSheetSelection.All())));
      WorkbookReadResult.NamedRangeHealthResult namedRangeHealth =
          cast(
              WorkbookReadResult.NamedRangeHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeNamedRangeHealth(
                      "namedRangeHealth", new ExcelNamedRangeSelection.All())));
      WorkbookReadResult.AutofilterHealthResult autofilterHealth =
          cast(
              WorkbookReadResult.AutofilterHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeAutofilterHealth(
                      "autofilterHealth", new ExcelSheetSelection.All())));
      WorkbookReadResult.ConditionalFormattingHealthResult conditionalFormattingHealth =
          cast(
              WorkbookReadResult.ConditionalFormattingHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
                      "conditionalFormattingHealth", new ExcelSheetSelection.All())));
      WorkbookReadResult.TableHealthResult tableHealth =
          cast(
              WorkbookReadResult.TableHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeTableHealth(
                      "tableHealth", new ExcelTableSelection.All())));
      WorkbookReadResult.WorkbookFindingsResult workbookFindings =
          cast(
              WorkbookReadResult.WorkbookFindingsResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeWorkbookFindings("workbookFindings")));

      assertEquals(2, formulaHealth.analysis().checkedFormulaCellCount());
      assertTrue(
          formulaHealth.analysis().findings().stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      WorkbookAnalysis.AnalysisFindingCode.FORMULA_ERROR_RESULT,
                      WorkbookAnalysis.AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION)));

      assertEquals(3, hyperlinkHealth.analysis().checkedHyperlinkCount());
      assertTrue(
          hyperlinkHealth.analysis().findings().stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET,
                      WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
                      WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET)));

      assertEquals(4, namedRangeHealth.analysis().checkedNamedRangeCount());
      assertTrue(
          namedRangeHealth.analysis().findings().stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
                      WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_UNRESOLVED_TARGET)));

      assertEquals(1, autofilterHealth.analysis().checkedAutofilterCount());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          autofilterHealth.analysis().findings().getFirst().code());

      assertEquals(
          1, conditionalFormattingHealth.analysis().checkedConditionalFormattingBlockCount());
      assertEquals(
          conditionalFormattingHealth.analysis().summary().totalCount(),
          conditionalFormattingHealth.analysis().findings().size());

      assertEquals(1, tableHealth.analysis().checkedTableCount());
      assertEquals(0, tableHealth.analysis().summary().totalCount());

      assertTrue(workbookFindings.analysis().summary().totalCount() >= 7);
      assertEquals(
          workbookFindings.analysis().summary().totalCount(),
          workbookFindings.analysis().findings().size());
      assertTrue(
          workbookFindings.analysis().findings().stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .contains(WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH));
    }
  }

  @Test
  void respectsSelectedSheetsAndNamedRanges() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.setCell("A1", ExcelCellValue.text("Item"));
      budget.setCell("B1", ExcelCellValue.formula("NOW()"));
      budget.setHyperlink("A1", new ExcelHyperlink.Document("Missing!A1"));

      ExcelSheet forecast = workbook.getOrCreateSheet("Forecast");
      forecast.setCell("A1", ExcelCellValue.text("Item"));
      forecast.setCell("B1", ExcelCellValue.formula("1+1"));
      forecast.setHyperlink("C1", new ExcelHyperlink.File("relative-report.xlsx"));

      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "B1")));

      WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

      WorkbookAnalysis.FormulaHealth formulaHealth =
          analyzer.formulaHealth(workbook, new ExcelSheetSelection.Selected(List.of("Forecast")));
      assertEquals(1, formulaHealth.checkedFormulaCellCount());
      assertEquals(0, formulaHealth.summary().infoCount());

      WorkbookAnalysis.HyperlinkHealth hyperlinkHealth =
          analyzer.hyperlinkHealth(workbook, new ExcelSheetSelection.Selected(List.of("Budget")));
      assertEquals(1, hyperlinkHealth.checkedHyperlinkCount());
      assertEquals(1, hyperlinkHealth.summary().errorCount());

      WorkbookAnalysis.HyperlinkHealth unresolvedFileHealth =
          analyzer.hyperlinkHealth(
              workbook,
              new WorkbookLocation.UnsavedWorkbook(),
              new ExcelSheetSelection.Selected(List.of("Forecast")));
      assertEquals(1, unresolvedFileHealth.checkedHyperlinkCount());
      assertEquals(1, unresolvedFileHealth.summary().warningCount());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET,
          unresolvedFileHealth.findings().getFirst().code());

      WorkbookAnalysis.NamedRangeHealth namedRangeHealth =
          analyzer.namedRangeHealth(
              workbook,
              new ExcelNamedRangeSelection.Selected(
                  List.of(new ExcelNamedRangeSelector.WorkbookScope("BudgetTotal"))));
      assertEquals(1, namedRangeHealth.checkedNamedRangeCount());
      assertEquals(0, namedRangeHealth.summary().totalCount());
    }
  }

  @Test
  void rejectsNullWorkbookAndSelections() {
    WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

    assertThrows(
        NullPointerException.class,
        () ->
            analyzer.execute(
                null,
                new WorkbookLocation.UnsavedWorkbook(),
                new WorkbookReadCommand.AnalyzeFormulaHealth(
                    "formulaHealth", new ExcelSheetSelection.All())));
    assertThrows(
        NullPointerException.class,
        () ->
            analyzer.execute(ExcelWorkbook.create(), new WorkbookLocation.UnsavedWorkbook(), null));
    assertThrows(
        NullPointerException.class,
        () ->
            analyzer.execute(
                ExcelWorkbook.create(),
                null,
                new WorkbookReadCommand.AnalyzeFormulaHealth(
                    "formulaHealth", new ExcelSheetSelection.All())));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.formulaHealth(null, new ExcelSheetSelection.All()));
    assertThrows(
        NullPointerException.class, () -> analyzer.formulaHealth(ExcelWorkbook.create(), null));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.hyperlinkHealth(null, new ExcelSheetSelection.All()));
    assertThrows(
        NullPointerException.class, () -> analyzer.hyperlinkHealth(ExcelWorkbook.create(), null));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.namedRangeHealth(null, new ExcelNamedRangeSelection.All()));
    assertThrows(
        NullPointerException.class, () -> analyzer.namedRangeHealth(ExcelWorkbook.create(), null));
  }

  @Test
  void workbookFindingsDefaultsToAnUnsavedWorkbookLocation() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook
          .getOrCreateSheet("Budget")
          .setHyperlink("A1", new ExcelHyperlink.File("reports/q1.xlsx"));

      WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

      WorkbookAnalysis.HyperlinkHealth hyperlinkHealth =
          analyzer.hyperlinkHealth(workbook, new ExcelSheetSelection.All());
      WorkbookAnalysis.WorkbookFindings workbookFindings = analyzer.workbookFindings(workbook);

      assertEquals(1, hyperlinkHealth.findings().size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET,
          hyperlinkHealth.findings().getFirst().code());
      assertTrue(
          workbookFindings.findings().stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .contains(WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET));
    }
  }

  @Test
  void scopeShadowingFindingsDetectWorkbookAndSheetScopeDuplicates() {
    List<ExcelNamedRangeSnapshot> namedRanges =
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "SharedName",
                new ExcelNamedRangeScope.WorkbookScope(),
                "Budget!$A$1",
                new ExcelNamedRangeTarget("Budget", "A1")),
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "SharedName",
                new ExcelNamedRangeScope.SheetScope("Budget"),
                "Budget!$B$2",
                new ExcelNamedRangeTarget("Budget", "B2")));

    List<WorkbookAnalysis.AnalysisFinding> findings =
        WorkbookAnalyzer.scopeShadowingFindings(namedRanges);

    assertEquals(2, findings.size());
    assertTrue(
        findings.stream()
            .allMatch(
                finding ->
                    finding.code()
                        == WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_SCOPE_SHADOWING));
  }

  @Test
  void scopeShadowingFindingsReturnsEmptyWhenScopesDoNotOverlap() {
    List<ExcelNamedRangeSnapshot> namedRanges =
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                "Budget!$A$1",
                new ExcelNamedRangeTarget("Budget", "A1")),
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "ForecastTotal",
                new ExcelNamedRangeScope.SheetScope("Forecast"),
                "Forecast!$B$2",
                new ExcelNamedRangeTarget("Forecast", "B2")));

    assertEquals(List.of(), WorkbookAnalyzer.scopeShadowingFindings(namedRanges));
  }

  @Test
  void namedRangeHealthHandlesQuotedLiteralAndUnresolvedFormulaShapes() throws Exception {
    Path workbookPath =
        assertDoesNotThrow(() -> Files.createTempFile("gridgrind-ranges-", ".xlsx"));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Quarter 1").createRow(0).createCell(0).setCellValue(42);
      workbook.createSheet("Budget").createRow(0).createCell(0).setCellValue(7);

      var quoted = workbook.createName();
      quoted.setNameName("QuarterOneCell");
      quoted.setRefersToFormula("'Quarter 1'!$A$1");

      var literal = workbook.createName();
      literal.setNameName("LiteralFormula");
      literal.setRefersToFormula("42");

      var unresolved = workbook.createName();
      unresolved.setNameName("UnresolvedLetterTail");
      unresolved.setRefersToFormula("Budget!A1+1");

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    } catch (IOException exception) {
      fail(exception);
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookAnalysis.NamedRangeHealth analysis =
          new WorkbookAnalyzer().namedRangeHealth(workbook, new ExcelNamedRangeSelection.All());

      assertEquals(3, analysis.checkedNamedRangeCount());
      assertEquals(1, analysis.summary().warningCount());
      assertEquals(0, analysis.summary().errorCount());
      assertEquals(
          List.of(WorkbookAnalysis.AnalysisFindingCode.NAMED_RANGE_UNRESOLVED_TARGET),
          analysis.findings().stream().map(WorkbookAnalysis.AnalysisFinding::code).toList());
    }
  }

  @Test
  void helperMethodsHandleQuotedLiteralAndDanglingReferenceShapes() throws Exception {
    WorkbookAnalyzer analyzer = new WorkbookAnalyzer();
    assertEquals("Quarter 1", analyzer.referencedSheetName("'Quarter 1'!$A$1"));
    assertNull(analyzer.referencedSheetName("42"));
    assertEquals("Budget", analyzer.referencedSheetName("Budget!A1+1"));
    assertEquals("'", analyzer.referencedSheetName("'!A1"));
    assertEquals("'Budget", analyzer.referencedSheetName("'Budget!A1"));

    assertTrue(analyzer.looksLikeSheetRangeReference("Budget!$A$1"));
    assertTrue(analyzer.looksLikeSheetRangeReference("Budget!A1:B2"));
    assertTrue(analyzer.looksLikeSheetRangeReference("Budget!A1+1"));
    assertFalse(analyzer.looksLikeSheetRangeReference("Budget!"));
    assertFalse(analyzer.looksLikeSheetRangeReference("Budget!1"));
    assertFalse(analyzer.looksLikeSheetRangeReference("42"));
  }

  @Test
  void summaryCountsAllSeverities() {
    WorkbookAnalyzer analyzer = new WorkbookAnalyzer();
    WorkbookAnalysis.AnalysisSummary result =
        analyzer.summary(
            List.of(
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    WorkbookAnalysis.AnalysisSeverity.ERROR,
                    "Error",
                    "Error finding",
                    new WorkbookAnalysis.AnalysisLocation.Workbook(),
                    List.of("1/0")),
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.FORMULA_EXTERNAL_REFERENCE,
                    WorkbookAnalysis.AnalysisSeverity.WARNING,
                    "Warning",
                    "Warning finding",
                    new WorkbookAnalysis.AnalysisLocation.Sheet("Budget"),
                    List.of("[Book]Sheet!A1")),
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                    WorkbookAnalysis.AnalysisSeverity.INFO,
                    "Info",
                    "Info finding",
                    new WorkbookAnalysis.AnalysisLocation.Cell("Budget", "A1"),
                    List.of("NOW()"))));

    assertEquals(new WorkbookAnalysis.AnalysisSummary(3, 1, 1, 1), result);
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }

  private static void writeAnalyzerFixture(Path workbookPath) throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var budget = workbook.createSheet("Budget");
      budget.createRow(0).createCell(0).setCellValue("Item");
      budget.getRow(0).createCell(1).setCellValue("Amount");
      budget.createRow(1).createCell(0).setCellValue("Hosting");
      budget.getRow(1).createCell(1).setCellFormula("1/0");
      budget.createRow(2).createCell(0).setCellValue("Domain");
      budget.getRow(2).createCell(1).setCellFormula("NOW()");

      namedRange(workbook, "SharedName", "Budget!$A$2");
      namedRange(workbook, "BrokenRange", "#REF!");
      namedRange(workbook, "MissingSheetRange", "Missing!$A$1");
      namedRange(workbook, "ApproximateRange", "Budget!$A$2+1");

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }
  }

  private static void namedRange(XSSFWorkbook workbook, String name, String formula) {
    var namedRange = workbook.createName();
    namedRange.setNameName(name);
    namedRange.setRefersToFormula(formula);
  }
}
