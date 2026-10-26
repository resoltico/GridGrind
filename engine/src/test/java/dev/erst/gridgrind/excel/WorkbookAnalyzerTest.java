package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.FormulaHealthReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeHealthReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.HyperlinkHealthReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookFindingsReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookAnalyzer finding-bearing analysis reads. */
class WorkbookAnalyzerTest {
  @Test
  void executesEveryAnalysisCommandAgainstWorkbookState() throws IOException {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-analyzer-", ".xlsx");
    writeAnalyzerFixture(workbookPath);

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet budget = workbook.sheet("Budget");
      budget.annotations().setHyperlink("A1", new HyperlinkTarget.Document("Missing!A1"));
      budget.annotations().setHyperlink("A2", new HyperlinkTarget.Document("Budget!A1:"));
      budget
          .annotations()
          .setHyperlink("A3", new HyperlinkTarget.File("missing-supporting-doc.pdf"));
      budget.cells().setCell("D1", new CellValue.Text("Owner"));
      budget.cells().setCell("E1", new CellValue.Text("Task"));
      budget.cells().setCell("D2", new CellValue.Text("Ada"));
      budget.cells().setCell("E2", new CellValue.Text("Queue"));
      budget.cells().setCell("D3", new CellValue.Text("Lin"));
      budget.cells().setCell("E3", new CellValue.Text("Pack"));
      budget
          .metadata()
          .setConditionalFormatting(
              List.of("B2:B3").stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                  .toList(),
              new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                  List.of(
                      new ConditionalFormattingRuleSpec.FormulaRule(
                          "B2>0",
                          false,
                          Optional.of(
                              new DifferentialStylePatch(
                                  Optional.of("0.00"),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty()))))));
      workbook
          .tables()
          .setTable(
              new ResolvedTable(
                  new TableName("Queue"),
                  new SheetName("Budget"),
                  new RangeRef("D1:E3"),
                  false,
                  new TableStyleSpec.None()));
      budget.xssfSheet().getTables().getFirst().getCTTable().getAutoFilter().setRef("D1:E2");
      workbook.getOrCreateSheet("PivotReport");
      workbook
          .pivots()
          .setPivotTable(
              new PivotTableSpec(
                  "Queue Pivot",
                  new dev.erst.gridgrind.core.vocabulary.SheetName("PivotReport"),
                  new PivotTableSpec.PivotSourceSpec.Range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("D1:E3")),
                  new PivotTableSpec.Anchor(
                      new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                  List.of("Owner"),
                  List.of(),
                  List.of(),
                  List.of(
                      new PivotTableSpec.DataField(
                          "Task",
                          ExcelPivotDataConsolidateFunction.COUNT,
                          "Task Count",
                          Optional.empty()))));
      workbook.xssfWorkbook().getPivotTables().getFirst().getCTPivotTableDefinition().setName(null);

      WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.FormulaHealthResult formulaHealth =
          cast(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.FormulaHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeFormulaHealth(
                      new StepId("formulaHealth"), new ExcelSheetSelection.All())));
      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.HyperlinkHealthResult hyperlinkHealth =
          cast(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.HyperlinkHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeHyperlinkHealth(
                      new StepId("hyperlinkHealth"), new ExcelSheetSelection.All())));
      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.NamedRangeHealthResult namedRangeHealth =
          cast(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.NamedRangeHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeNamedRangeHealth(
                      new StepId("namedRangeHealth"), new ExcelNamedRangeSelection.All())));
      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.AutofilterHealthResult autofilterHealth =
          cast(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.AutofilterHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeAutofilterHealth(
                      new StepId("autofilterHealth"), new ExcelSheetSelection.All())));
      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.ConditionalFormattingHealthResult
          conditionalFormattingHealth =
              cast(
                  dev.erst.gridgrind.core.plan.WorkbookAnalysisResult
                      .ConditionalFormattingHealthResult.class,
                  analyzer.execute(
                      workbook,
                      new WorkbookLocation.StoredWorkbook(workbookPath),
                      new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
                          new StepId("conditionalFormattingHealth"),
                          new ExcelSheetSelection.All())));
      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.TableHealthResult tableHealth =
          cast(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.TableHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeTableHealth(
                      new StepId("tableHealth"), new ExcelTableSelection.All())));
      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.PivotTableHealthResult pivotTableHealth =
          cast(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.PivotTableHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzePivotTableHealth(
                      new StepId("pivotTableHealth"), new ExcelPivotTableSelection.All())));
      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.WorkbookFindingsResult workbookFindings =
          cast(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.WorkbookFindingsResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeWorkbookFindings(new StepId("workbookFindings"))));

      assertEquals(2, formulaHealth.analysis().checkedFormulaCellCount());
      assertTrue(
          formulaHealth.analysis().findings().stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.FORMULA_ERROR_RESULT,
                      AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION)));

      assertEquals(3, hyperlinkHealth.analysis().checkedHyperlinkCount());
      assertTrue(
          hyperlinkHealth.analysis().findings().stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET,
                      AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
                      AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET)));

      assertEquals(4, namedRangeHealth.analysis().checkedNamedRangeCount());
      assertTrue(
          namedRangeHealth.analysis().findings().stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .containsAll(List.of(AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE)));

      assertEquals(1, autofilterHealth.analysis().checkedAutofilterCount());
      assertEquals(
          AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          autofilterHealth.analysis().findings().getFirst().code());

      assertEquals(
          1, conditionalFormattingHealth.analysis().checkedConditionalFormattingBlockCount());
      assertEquals(
          conditionalFormattingHealth.analysis().summary().totalCount(),
          conditionalFormattingHealth.analysis().findings().size());

      assertEquals(1, tableHealth.analysis().checkedTableCount());
      assertEquals(0, tableHealth.analysis().summary().totalCount());

      assertEquals(1, pivotTableHealth.analysis().checkedPivotTableCount());
      assertEquals(
          List.of(AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME),
          pivotTableHealth.analysis().findings().stream()
              .map(AnalysisFindingReport::code)
              .distinct()
              .toList());

      assertTrue(workbookFindings.analysis().summary().totalCount() >= 8);
      assertEquals(
          workbookFindings.analysis().summary().totalCount(),
          workbookFindings.analysis().findings().size());
      assertTrue(
          workbookFindings.analysis().findings().stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .contains(AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH));
      assertTrue(
          workbookFindings.analysis().findings().stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .contains(AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME));
    }
  }

  @Test
  void respectsSelectedSheetsAndNamedRanges() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.cells().setCell("A1", new CellValue.Text("Item"));
      budget.cells().setCell("B1", new ResolvedFormula("NOW()"));
      budget.annotations().setHyperlink("A1", new HyperlinkTarget.Document("Missing!A1"));

      ExcelSheet forecast = workbook.getOrCreateSheet("Forecast");
      forecast.cells().setCell("A1", new CellValue.Text("Item"));
      forecast.cells().setCell("B1", new ResolvedFormula("1+1"));
      forecast.annotations().setHyperlink("C1", new HyperlinkTarget.File("relative-report.xlsx"));

      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("B1"))));

      WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

      FormulaHealthReport formulaHealth =
          analyzer.formulaHealth(workbook, new ExcelSheetSelection.Selected(List.of("Forecast")));
      assertEquals(1, formulaHealth.checkedFormulaCellCount());
      assertEquals(0, formulaHealth.summary().infoCount());

      HyperlinkHealthReport hyperlinkHealth =
          analyzer.hyperlinkHealth(workbook, new ExcelSheetSelection.Selected(List.of("Budget")));
      assertEquals(1, hyperlinkHealth.checkedHyperlinkCount());
      assertEquals(1, hyperlinkHealth.summary().errorCount());

      HyperlinkHealthReport unresolvedFileHealth =
          analyzer.hyperlinkHealth(
              workbook,
              new WorkbookLocation.UnsavedWorkbook(),
              new ExcelSheetSelection.Selected(List.of("Forecast")));
      assertEquals(1, unresolvedFileHealth.checkedHyperlinkCount());
      assertEquals(1, unresolvedFileHealth.summary().warningCount());
      assertEquals(
          AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET,
          unresolvedFileHealth.findings().getFirst().code());

      NamedRangeHealthReport namedRangeHealth =
          analyzer.namedRangeHealth(
              workbook,
              new ExcelNamedRangeSelection.Selected(
                  List.of(new ExcelNamedRangeSelector.WorkbookScope("BudgetTotal"))));
      assertEquals(1, namedRangeHealth.checkedNamedRangeCount());
      assertEquals(
          0,
          namedRangeHealth.summary().totalCount(),
          () -> "unexpected named-range findings: " + namedRangeHealth.findings());
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
                    new StepId("formulaHealth"), new ExcelSheetSelection.All())));
    assertThrows(
        NullPointerException.class,
        () ->
            analyzer.execute(
                ExcelWorkbooks.create(), new WorkbookLocation.UnsavedWorkbook(), null));
    assertThrows(
        NullPointerException.class,
        () ->
            analyzer.execute(
                ExcelWorkbooks.create(),
                null,
                new WorkbookReadCommand.AnalyzeFormulaHealth(
                    new StepId("formulaHealth"), new ExcelSheetSelection.All())));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.formulaHealth(null, new ExcelSheetSelection.All()));
    assertThrows(
        NullPointerException.class, () -> analyzer.formulaHealth(ExcelWorkbooks.create(), null));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.hyperlinkHealth(null, new ExcelSheetSelection.All()));
    assertThrows(
        NullPointerException.class, () -> analyzer.hyperlinkHealth(ExcelWorkbooks.create(), null));
    assertThrows(
        NullPointerException.class,
        () -> analyzer.namedRangeHealth(null, new ExcelNamedRangeSelection.All()));
    assertThrows(
        NullPointerException.class, () -> analyzer.namedRangeHealth(ExcelWorkbooks.create(), null));
  }

  @Test
  void workbookFindingsDefaultsToAnUnsavedWorkbookLocation() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Budget")
          .annotations()
          .setHyperlink("A1", new HyperlinkTarget.File("reports/q1.xlsx"));

      WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

      HyperlinkHealthReport hyperlinkHealth =
          analyzer.hyperlinkHealth(workbook, new ExcelSheetSelection.All());
      WorkbookFindingsReport workbookFindings = analyzer.workbookFindings(workbook);

      assertEquals(1, hyperlinkHealth.findings().size());
      assertEquals(
          AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET,
          hyperlinkHealth.findings().getFirst().code());
      assertTrue(
          workbookFindings.findings().stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .contains(AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET));
    }
  }

  @Test
  void scopeShadowingFindingsDetectWorkbookAndSheetScopeDuplicates() {
    List<NamedRangeReport> namedRanges =
        List.of(
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("SharedName"),
                new NamedRangeScopeReport.Workbook(),
                "Budget!$A$1",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"))),
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("SharedName"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
                "Budget!$B$2",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B2"))));

    List<AnalysisFindingReport> findings = WorkbookAnalyzer.scopeShadowingFindings(namedRanges);

    assertEquals(2, findings.size());
    assertTrue(
        findings.stream()
            .allMatch(
                finding -> finding.code() == AnalysisFindingCode.NAMED_RANGE_SCOPE_SHADOWING));
  }

  @Test
  void scopeShadowingFindingsReturnsEmptyWhenScopesDoNotOverlap() {
    List<NamedRangeReport> namedRanges =
        List.of(
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
                new NamedRangeScopeReport.Workbook(),
                "Budget!$A$1",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1"))),
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("ForecastTotal"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Forecast")),
                "Forecast!$B$2",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Forecast"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B2"))));

    assertEquals(List.of(), WorkbookAnalyzer.scopeShadowingFindings(namedRanges));
  }

  @Test
  void namedRangeHealthHandlesQuotedLiteralAndUnresolvedFormulaShapes() throws Exception {
    Path workbookPath =
        assertDoesNotThrow(
            () -> ExcelTempFiles.createManagedTempFile("gridgrind-ranges-", ".xlsx"));

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

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      NamedRangeHealthReport analysis =
          new WorkbookAnalyzer().namedRangeHealth(workbook, new ExcelNamedRangeSelection.All());

      assertEquals(3, analysis.checkedNamedRangeCount());
      assertEquals(0, analysis.summary().warningCount());
      assertEquals(0, analysis.summary().errorCount());
      assertTrue(analysis.findings().isEmpty());
    }
  }

  @Test
  void formulaNamedRangesUseWorkbookContextInsteadOfTextHeuristics() throws Exception {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-named-formulas-", ".xlsx");
    Path patchedWorkbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-named-formulas-patched-", ".xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Quarter 1");
      workbook.createSheet("O'Brien");
      workbook.createSheet("Budget");
      var literal = workbook.createName();
      literal.setNameName("LiteralFormula");
      literal.setRefersToFormula("42");
      var quoted = workbook.createName();
      quoted.setNameName("QuotedSheet");
      quoted.setRefersToFormula("'Quarter 1'!$A$1");
      var apostrophe = workbook.createName();
      apostrophe.setNameName("ApostropheSheet");
      apostrophe.setRefersToFormula("'O''Brien'!$A$1");
      var multiple = workbook.createName();
      multiple.setNameName("MultipleSheets");
      multiple.setRefersToFormula("Budget!$A$1+'Quarter 1'!$A$1");
      var threeDimensional = workbook.createName();
      threeDimensional.setNameName("ThreeDimensionalSheets");
      threeDimensional.setRefersToFormula("Budget:'Quarter 1'!$A$1");
      var local = workbook.createName();
      local.setNameName("LocalFormula");
      local.setSheetIndex(workbook.getSheetIndex("Budget"));
      local.setRefersToFormula("Budget!$A$1");
      var localNameReference = workbook.createName();
      localNameReference.setNameName("LocalNameReference");
      localNameReference.setRefersToFormula("Budget!LocalFormula");
      try (XSSFWorkbook externalWorkbook = new XSSFWorkbook()) {
        externalWorkbook.createSheet("Ledger");
        workbook.linkExternalWorkbook("External.xlsx", externalWorkbook);
      }
      var external = workbook.createName();
      external.setNameName("ExternalFormula");
      external.setRefersToFormula("[External.xlsx]Ledger!$A$1");
      var missing = workbook.createName();
      missing.setNameName("MissingSheet");
      missing.setRefersToFormula("Missing!$A$1+Missing!$B$1");
      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    appendUnparseableDefinedName(workbookPath, patchedWorkbookPath);

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            patchedWorkbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {

      NamedRangeHealthReport analysis =
          new WorkbookAnalyzer().namedRangeHealth(workbook, new ExcelNamedRangeSelection.All());

      assertEquals(10, analysis.checkedNamedRangeCount());
      assertEquals(
          List.of(
              AnalysisFindingCode.NAMED_RANGE_BROKEN_REFERENCE,
              AnalysisFindingCode.NAMED_RANGE_UNPARSEABLE_FORMULA),
          analysis.findings().stream().map(AnalysisFindingReport::code).toList());
      assertEquals(
          "Named range refers to a sheet that does not exist: Missing",
          analysis.findings().getFirst().message());
    }
  }

  private static void appendUnparseableDefinedName(Path sourcePath, Path targetPath)
      throws IOException {
    try (ZipFile source = new ZipFile(sourcePath.toFile());
        ZipOutputStream target = new ZipOutputStream(Files.newOutputStream(targetPath))) {
      for (ZipEntry entry : source.stream().toList()) {
        copyPatchedDefinedNameEntry(source, target, entry);
      }
    }
  }

  private static void copyPatchedDefinedNameEntry(
      ZipFile source, ZipOutputStream target, ZipEntry entry) throws IOException {
    target.putNextEntry(new ZipEntry(entry.getName()));
    try (var input = source.getInputStream(entry)) {
      if ("xl/workbook.xml".equals(entry.getName())) {
        String xml = new String(input.readAllBytes(), StandardCharsets.UTF_8);
        String patchedXml =
            xml.replace(
                "</definedNames>",
                "<definedName name=\"UnparseableFormula\">SUM(</definedName></definedNames>");
        if (patchedXml.equals(xml)) {
          throw new IllegalStateException("Expected workbook.xml to contain defined names");
        }
        target.write(patchedXml.getBytes(StandardCharsets.UTF_8));
      } else {
        input.transferTo(target);
      }
    }
    target.closeEntry();
  }

  @Test
  void summaryCountsAllSeverities() {
    WorkbookAnalyzer analyzer = new WorkbookAnalyzer();
    AnalysisSummaryReport result =
        analyzer.summary(
            List.of(
                new AnalysisFindingReport(
                    AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    AnalysisSeverity.ERROR,
                    "Error",
                    "Error finding",
                    new AnalysisLocationReport.Workbook(),
                    List.of("1/0")),
                new AnalysisFindingReport(
                    AnalysisFindingCode.FORMULA_EXTERNAL_REFERENCE,
                    AnalysisSeverity.WARNING,
                    "Warning",
                    "Warning finding",
                    new AnalysisLocationReport.Sheet(
                        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
                    List.of("[Book]Sheet!A1")),
                new AnalysisFindingReport(
                    AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                    AnalysisSeverity.INFO,
                    "Info",
                    "Info finding",
                    new AnalysisLocationReport.Cell(
                        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"), "A1"),
                    List.of("NOW()"))));

    assertEquals(new AnalysisSummaryReport(3, 1, 1, 1), result);
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
