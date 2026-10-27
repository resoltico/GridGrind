package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientStopSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
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
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-engine-", ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine commandExecutor = new WorkbookExecutionEngine();
      commandExecutor.apply(
          workbook,
          List.<ResolvedStep>of(
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new ResolvedMutationAction.AppendRow(
                      List.of(new CellValue.Text("Item"), new CellValue.Text("Amount")))),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new ResolvedMutationAction.AppendRow(
                      List.of(new CellValue.Text("Hosting"), new CellValue.NumberValue(49.0)))),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new ResolvedMutationAction.AppendRow(
                      List.of(new CellValue.Text("Domain"), new CellValue.NumberValue(12.0)))),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A4")),
                  new ResolvedMutationAction.SetCell(new CellValue.Text("Total"))),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B4")),
                  new ResolvedMutationAction.SetCell(new ResolvedFormula("SUM(B2:B3)")))));
      workbook.formulas().evaluateAll();
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelSheet sheet = workbook.sheet("Budget");

      CellReport.FormulaReport totalSnapshot =
          (CellReport.FormulaReport) sheet.cells().snapshotCell("B4");
      assertEquals("FORMULA", totalSnapshot.type());
      assertEquals("FORMULA", totalSnapshot.type());
      assertEquals("SUM(B2:B3)", totalSnapshot.formula().orElseThrow());
      assertEquals(
          61.0,
          ((dev.erst.gridgrind.core.model.value.CellValue.NumberValue)
                  totalSnapshot.value().orElseThrow())
              .value());

      List<ExcelPreviewRow> preview = sheet.cells().preview(4, 2);
      assertEquals(4, preview.size());
      assertEquals("A1", preview.get(0).cells().get(0).address().value());
      assertEquals(
          "Hosting",
          ((CellReport.TextReport) preview.get(1).cells().get(0)).value().orElseThrow().value());
      assertEquals("61", preview.get(3).cells().get(1).displayValue().orElseThrow());
    }
  }

  @Test
  void managesWorkbookLifecycleAndValidation() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempDirectory("gridgrind-workbook-")
            .resolve("nested/book.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Hello"));
      workbook.getOrCreateSheet("Budget").cells().setCell("B1", new CellValue.NumberValue(12.0));
      workbook.formulas().markRecalculateOnOpen();
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());

      assertEquals(1, workbook.sheets().sheetCount());
      assertEquals(List.of("Budget"), workbook.sheets().sheetNames());
      assertTrue(workbook.formulas().recalculateOnOpenEnabled());
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Hello", workbook.sheet("Budget").cells().text("A1"));
      assertEquals(12.0, workbook.sheet("Budget").cells().number("B1"));
    }
  }

  @Test
  void createAndOpenWithFormulaEnvironmentExposeRuntimeContextAndValidateFormulaTargets()
      throws IOException {
    Path directory = ExcelTempFiles.createManagedTempDirectory("gridgrind-formula-workbook-");
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

    try (ExcelWorkbook workbook = ExcelWorkbooks.create(environment)) {
      assertTrue(workbook.formulaRuntimeContext().hasExternalWorkbookBinding("RATES.XLSX"));
      assertTrue(workbook.formulaRuntimeContext().hasUserDefinedFunction("double"));
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", new CellValue.NumberValue(2.0d));
      workbook.sheet("Budget").cells().setCell("B1", new ResolvedFormula("A1*2"));
      workbook.formulas().evaluateAll();
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());

      assertThrows(
          IllegalArgumentException.class,
          () -> workbook.formulas().evaluate(List.of(new ExcelFormulaCellTarget("Budget", "A1"))));
      assertThrows(
          InvalidCellAddressException.class,
          () -> workbook.formulas().evaluate(List.of(new ExcelFormulaCellTarget("Budget", ":"))));
      assertThrows(
          CellNotFoundException.class,
          () -> workbook.formulas().evaluate(List.of(new ExcelFormulaCellTarget("Budget", "Z99"))));
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(
            workbookPath, environment, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertTrue(reopened.formulaRuntimeContext().hasExternalWorkbookBinding("rates.xlsx"));
      assertTrue(reopened.formulaRuntimeContext().hasUserDefinedFunction("DOUBLE"));
    }
  }

  @Test
  void renamesDeletesAndMovesSheetsAcrossSaves() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-sheet-management-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Live"));
      workbook.getOrCreateSheet("Archive").cells().setCell("A1", new CellValue.Text("Old"));
      workbook.getOrCreateSheet("Scratch");

      workbook.sheets().renameSheet("Archive", "History");
      workbook.sheets().moveSheet("History", 0);
      workbook.sheets().deleteSheet("Scratch");
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());

      assertEquals(List.of("History", "Budget"), workbook.sheets().sheetNames());
      assertEquals("Old", workbook.sheet("History").cells().text("A1"));
      assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Archive"));
    }

    assertEquals(List.of("History", "Budget"), XlsxRoundTrip.sheetOrder(workbookPath));
  }

  @Test
  void workbookSummaryUsesExplicitEmptyAndWithSheetsStates() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookSummary emptySummary = workbook.workbookSummary();
      WorkbookSummary.Empty empty = assertInstanceOf(WorkbookSummary.Empty.class, emptySummary);
      assertEquals(0, empty.sheetCount());
      assertEquals(List.of(), empty.sheetNames());

      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.getOrCreateSheet("Gamma");
      workbook.sheets().setActiveSheet("Beta");
      workbook.sheets().setSelectedSheets(List.of("Gamma", "Alpha"));
      workbook.sheets().setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN);

      WorkbookSummary.WithSheets summary =
          assertInstanceOf(WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals(
          List.of("Alpha", "Beta", "Gamma").stream()
              .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
              .toList(),
          summary.sheetNames());
      assertEquals(
          new dev.erst.gridgrind.core.vocabulary.SheetLabel("Gamma"), summary.activeSheetName());
      assertEquals(
          List.of("Alpha", "Gamma").stream()
              .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
              .toList(),
          summary.selectedSheetNames());
    }
  }

  @Test
  void sheetStateRoundTripsAcrossSaveAndReopen() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-sheet-state-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha").cells().setCell("A1", new CellValue.Text("Live"));
      workbook.getOrCreateSheet("Beta");
      workbook.getOrCreateSheet("Gamma");
      workbook.sheets().copySheet("Alpha", "Replica", new SheetCopyPosition.AtIndex(1));
      workbook.sheets().setActiveSheet("Beta");
      workbook.sheets().setSelectedSheets(List.of("Gamma", "Alpha"));
      workbook.sheets().setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN);
      workbook.sheets().setSheetVisibility("Replica", ExcelSheetVisibility.VERY_HIDDEN);
      workbook.sheets().setSheetProtection("Alpha", protectionSettings());
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(
        List.of("Alpha", "Replica", "Beta", "Gamma"), XlsxRoundTrip.sheetOrder(workbookPath));
    assertEquals("Gamma", XlsxRoundTrip.activeSheetName(workbookPath));
    assertEquals(List.of("Alpha", "Gamma"), XlsxRoundTrip.selectedSheetNames(workbookPath));
    assertEquals(ExcelSheetVisibility.HIDDEN, XlsxRoundTrip.sheetVisibility(workbookPath, "Beta"));
    assertEquals(
        ExcelSheetVisibility.VERY_HIDDEN, XlsxRoundTrip.sheetVisibility(workbookPath, "Replica"));
    assertEquals(
        new SheetProtectionReport.Protected(protectionSettings()),
        XlsxRoundTrip.sheetProtection(workbookPath, "Alpha"));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      WorkbookSummary.WithSheets summary =
          assertInstanceOf(WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals(
          new dev.erst.gridgrind.core.vocabulary.SheetLabel("Gamma"), summary.activeSheetName());
      assertEquals(
          List.of("Alpha", "Gamma").stream()
              .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
              .toList(),
          summary.selectedSheetNames());

      SheetSummaryReport alphaSummary = workbook.sheetSummary("Alpha");
      SheetSummaryReport betaSummary = workbook.sheetSummary("Beta");
      SheetSummaryReport replicaSummary = workbook.sheetSummary("Replica");
      assertEquals(ExcelSheetVisibility.VISIBLE, alphaSummary.visibility());
      assertEquals(
          new SheetProtectionReport.Protected(protectionSettings()), alphaSummary.protection());
      assertEquals(ExcelSheetVisibility.HIDDEN, betaSummary.visibility());
      assertInstanceOf(SheetProtectionReport.Unprotected.class, betaSummary.protection());
      assertEquals(ExcelSheetVisibility.VERY_HIDDEN, replicaSummary.visibility());
    }
  }

  @Test
  void clearSheetProtectionIsIdempotentForUnprotectedSheets() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-clear-sheet-protection-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.sheets().clearSheetProtection("Alpha");
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertInstanceOf(
          SheetProtectionReport.Unprotected.class, workbook.sheetSummary("Alpha").protection());
    }
  }

  @Test
  void clearSheetProtectionRemovesExistingProtectionAcrossSaveAndReopen() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-clear-existing-sheet-protection-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.sheets().setSheetProtection("Alpha", protectionSettings());
      workbook.sheets().clearSheetProtection("Alpha");
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(
        new SheetProtectionReport.Unprotected(),
        XlsxRoundTrip.sheetProtection(workbookPath, "Alpha"));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertInstanceOf(
          SheetProtectionReport.Unprotected.class, workbook.sheetSummary("Alpha").protection());
    }
  }

  @Test
  void copySheetPreservesSupportedLocalStructuresAndCopiesSheetScopedRangeNames()
      throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-sheet-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Source");
      workbook.sheet("Source").cells().setCell("A1", new CellValue.Text("Item"));
      workbook.sheet("Source").cells().setCell("B1", new CellValue.Text("Amount"));
      workbook.sheet("Source").cells().setCell("A2", new CellValue.Text("Hosting"));
      workbook.sheet("Source").cells().setCell("B2", new CellValue.NumberValue(49.0));
      workbook.sheet("Source").cells().setCell("B3", new ResolvedFormula("SUM(B2:B2)"));
      workbook
          .sheet("Source")
          .annotations()
          .setHyperlink("A2", new HyperlinkTarget.Url("https://example.com/h"));
      workbook
          .sheet("Source")
          .annotations()
          .setComment("A2", new ResolvedComment("Review", "GridGrind", false));
      workbook
          .sheet("Source")
          .metadata()
          .setDataValidation(
              "C2:C4",
              new ResolvedDataValidation(
                  new DataValidationRuleSpec.WholeNumber(
                      ExcelComparisonOperator.GREATER_OR_EQUAL,
                      ExcelComparisonFormulaSupport.normalizeFormula("1", "formula1"),
                      ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                          Optional.empty(), "formula2")),
                  false,
                  false,
                  Optional.empty(),
                  Optional.empty()));
      workbook
          .sheet("Source")
          .metadata()
          .setConditionalFormatting(
              List.of("B2:B4").stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                  .toList(),
              new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                  List.of(
                      new ConditionalFormattingRuleSpec.FormulaRule(
                          "B2>0",
                          true,
                          Optional.of(
                              new DifferentialStylePatch(
                                  Optional.of("0.00"),
                                  Optional.of(true),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.of(ColorSpec.rgb(new RgbHex("#102030"))),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.of(ColorSpec.rgb(new RgbHex("#E0F0AA"))),
                                  Optional.empty()))))));
      workbook.sheet("Source").layout().mergeCells("A1:B1");
      workbook.sheet("Source").layout().setPane(new PaneSpec.Frozen(1, 1, 1, 1));
      workbook.sheet("Source").layout().setZoom(140);
      workbook
          .sheet("Source")
          .layout()
          .setPrintLayout(
              new ResolvedPrintLayout(
                  new PrintAreaSpec.Range(new RangeRef("A1:C20")),
                  ExcelPrintOrientation.LANDSCAPE,
                  new PrintScalingSpec.Fit(1, 0),
                  new PrintTitleRowsSpec.Band(0, 0),
                  new PrintTitleColumnsSpec.None(),
                  new HeaderFooterTextReport("Source", "", ""),
                  new HeaderFooterTextReport("", "&P", "")));
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("LocalBudget"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Source")),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Source"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B3"))));
      workbook.sheets().copySheet("Source", "Replica", new SheetCopyPosition.AppendAtEnd());
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(
        new dev.erst.gridgrind.core.model.layout.PaneReport.Frozen(1, 1, 1, 1),
        XlsxRoundTrip.pane(workbookPath, "Replica"));
    assertEquals(140, XlsxRoundTrip.zoomPercent(workbookPath, "Replica"));
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        XlsxRoundTrip.printLayout(workbookPath, "Replica").orientation());
    assertEquals(
        List.of(
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalBudget"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Source")),
                "Source!$A$1:$B$3",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Source"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B3"))),
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalBudget"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Replica")),
                "Replica!$A$1:$B$3",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Replica"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B3")))),
        XlsxRoundTrip.namedRanges(workbookPath));
    assertEquals(
        List.of(
            new DataValidationEntryReport.Supported(
                List.of("C2:C4").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                    .toList(),
                dev.erst.gridgrind.excel.DataValidationFixtures.report(
                    new ResolvedDataValidation(
                        new DataValidationRuleSpec.WholeNumber(
                            ExcelComparisonOperator.GREATER_OR_EQUAL,
                            ExcelComparisonFormulaSupport.normalizeFormula("1", "formula1"),
                            ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                                Optional.empty(), "formula2")),
                        false,
                        true,
                        Optional.empty(),
                        Optional.empty())))),
        XlsxRoundTrip.dataValidations(workbookPath, "Replica"));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Item", workbook.sheet("Replica").cells().text("A1"));
      assertEquals(
          "https://example.com/h",
          workbook
              .sheet("Replica")
              .cells()
              .snapshotCell("A2")
              .hyperlink()
              .map(link -> assertInstanceOf(HyperlinkTarget.Url.class, link).target())
              .orElseThrow());
      assertEquals(
          "Review",
          workbook.sheet("Replica").cells().snapshotCell("A2").comment().orElseThrow().text());
      assertEquals(
          List.of("A1:B1"),
          workbook.sheet("Replica").layout().mergedRegions().stream()
              .map(region -> region.range().value())
              .toList());
      assertEquals(
          List.of("B2:B4"),
          workbook.sheet("Replica").metadata().conditionalFormatting().getFirst().ranges().stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
              .toList());
    }
  }

  @Test
  void copySheetRejectsUnsupportedSourceStructuresAndVisibilityRulesStayHonest()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.sheets().setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN);
      assertDoesNotThrow(
          () -> workbook.sheets().setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN));

      IllegalArgumentException lastVisible =
          assertThrows(
              IllegalArgumentException.class,
              () -> workbook.sheets().setSheetVisibility("Alpha", ExcelSheetVisibility.HIDDEN));
      assertEquals("cannot hide the last visible sheet 'Alpha'", lastVisible.getMessage());

      IllegalArgumentException deleteLastVisible =
          assertThrows(
              IllegalArgumentException.class, () -> workbook.sheets().deleteSheet("Alpha"));
      assertEquals("cannot delete the last visible sheet 'Alpha'", deleteLastVisible.getMessage());
    }

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Tables");
      workbook.sheet("Tables").cells().setCell("A1", new CellValue.Text("Name"));
      workbook.sheet("Tables").cells().setCell("B1", new CellValue.Text("Value"));
      workbook.sheet("Tables").cells().setCell("A2", new CellValue.Text("Ops"));
      workbook.sheet("Tables").cells().setCell("B2", new CellValue.NumberValue(1.0));
      workbook
          .tables()
          .setTable(
              new ResolvedTable(
                  new TableName("OpsTable"),
                  new SheetName("Tables"),
                  new RangeRef("A1:B2"),
                  false,
                  new TableStyleSpec.None()));

      workbook.sheets().copySheet("Tables", "Tables Copy", new SheetCopyPosition.AppendAtEnd());

      assertEquals(List.of("Tables", "Tables Copy"), workbook.sheets().sheetNames());
      assertEquals(1, workbook.xssfWorkbook().getSheet("Tables").getTables().size());
      assertEquals(1, workbook.xssfWorkbook().getSheet("Tables Copy").getTables().size());
      assertEquals(
          "OpsTable", workbook.xssfWorkbook().getSheet("Tables").getTables().getFirst().getName());
      assertEquals(
          "OpsTable_Copy2",
          workbook.xssfWorkbook().getSheet("Tables Copy").getTables().getFirst().getName());
    }

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("FormulaNames");
      Name localFormula = workbook.xssfWorkbook().createName();
      localFormula.setNameName("LocalFormula");
      localFormula.setSheetIndex(workbook.xssfWorkbook().getSheetIndex("FormulaNames"));
      localFormula.setRefersToFormula("SUM(FormulaNames!$A$1:$A$2)");

      workbook
          .sheets()
          .copySheet("FormulaNames", "FormulaNames Copy", new SheetCopyPosition.AppendAtEnd());

      assertEquals(
          List.of(
              new NamedRangeReport.FormulaReport(
                  new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalFormula"),
                  new NamedRangeScopeReport.Sheet(
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel("FormulaNames")),
                  "SUM(FormulaNames!$A$1:$A$2)"),
              new NamedRangeReport.FormulaReport(
                  new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalFormula"),
                  new NamedRangeScopeReport.Sheet(
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel("FormulaNames Copy")),
                  "SUM('FormulaNames Copy'!$A$1:$A$2)")),
          workbook.names().namedRanges());
    }
  }

  @Test
  void persistsStructuralLayoutOperationsAcrossSaves() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-layout-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.cells().setCell("A1", new CellValue.Text("Quarterly"));
      budget.layout().mergeCells("A1:B1");
      budget.columns().setWidth(0, 1, 16.0);
      budget.rows().setHeight(0, 0, 28.5);
      budget.layout().setPane(new PaneSpec.Frozen(1, 1, 1, 1));
      budget.layout().setZoom(125);
      budget
          .layout()
          .setPrintLayout(
              new ResolvedPrintLayout(
                  new PrintAreaSpec.Range(new RangeRef("A1:B12")),
                  ExcelPrintOrientation.LANDSCAPE,
                  new PrintScalingSpec.Fit(1, 0),
                  new PrintTitleRowsSpec.Band(0, 0),
                  new PrintTitleColumnsSpec.Band(0, 0),
                  new HeaderFooterTextReport("Budget", "", ""),
                  new HeaderFooterTextReport("", "Page &P", "")));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(List.of("A1:B1"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(
        new dev.erst.gridgrind.core.model.layout.PaneReport.Frozen(1, 1, 1, 1),
        XlsxRoundTrip.pane(workbookPath, "Budget"));
    assertEquals(125, XlsxRoundTrip.zoomPercent(workbookPath, "Budget"));
    assertEquals(
        new PrintAreaSpec.Range(new RangeRef("A1:B12")),
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
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("1+1"));

      InvalidFormulaException exception =
          assertThrows(InvalidFormulaException.class, workbook.formulas()::evaluateAll);
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
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("1+1"));

      workbook.formulas().evaluateAll();

      CellReport.FormulaReport snapshot =
          (CellReport.FormulaReport) workbook.sheet("Budget").cells().snapshotCell("A1");
      assertEquals("2", snapshot.displayValue().orElseThrow());
    }
  }

  @Test
  void setCellLiteralReplacesExistingFormulaCell() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("1+1"));

      workbook.sheet("Budget").cells().setCell("A1", new CellValue.NumberValue(0.0d));

      CellReport.NumberReport snapshot =
          assertInstanceOf(
              CellReport.NumberReport.class, workbook.sheet("Budget").cells().snapshotCell("A1"));
      assertEquals("NUMBER", snapshot.type());
      assertEquals(0.0d, snapshot.value().orElseThrow().value());
    }
  }

  @Test
  void setRangeLiteralReplacesExistingFormulaCells() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("1+1"));
      workbook.getOrCreateSheet("Budget").cells().setCell("B1", new ResolvedFormula("2+2"));

      workbook
          .sheet("Budget")
          .cells()
          .setRange(
              "A1:B1",
              List.of(List.of(new CellValue.NumberValue(0.0d), new CellValue.NumberValue(5.0d))));

      CellReport.NumberReport first =
          assertInstanceOf(
              CellReport.NumberReport.class, workbook.sheet("Budget").cells().snapshotCell("A1"));
      CellReport.NumberReport second =
          assertInstanceOf(
              CellReport.NumberReport.class, workbook.sheet("Budget").cells().snapshotCell("B1"));
      assertEquals(0.0d, first.value().orElseThrow().value());
      assertEquals(5.0d, second.value().orElseThrow().value());
    }
  }

  @Test
  void validatesWorkbookInputsAndMissingResources() throws IOException {
    assertThrows(
        NullPointerException.class,
        () -> ExcelWorkbooks.open(null, ExcelTempFileFactoryTestSupport.tempFileFactory()));

    Path missingPath =
        ExcelTempFiles.createManagedTempDirectory("gridgrind-missing-").resolve("missing.xlsx");
    WorkbookNotFoundException missingWorkbook =
        assertThrows(
            WorkbookNotFoundException.class,
            () ->
                ExcelWorkbooks.open(
                    missingPath, ExcelTempFileFactoryTestSupport.tempFileFactory()));
    assertTrue(missingWorkbook.getMessage().contains("Workbook does not exist"));
    assertEquals(missingPath.toAbsolutePath(), missingWorkbook.workbookPath());

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelWorkbookSheets sheets = workbook.sheets();
      assertThrows(NullPointerException.class, () -> workbook.getOrCreateSheet(null));
      assertThrows(IllegalArgumentException.class, () -> workbook.getOrCreateSheet(" "));
      assertThrows(NullPointerException.class, () -> workbook.sheet(null));
      assertThrows(IllegalArgumentException.class, () -> workbook.sheet(" "));
      SheetNotFoundException missingSheet =
          assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Missing"));
      assertEquals("Missing", missingSheet.sheetName());
      workbook.getOrCreateSheet("Budget");
      workbook.getOrCreateSheet("Archive");
      assertSame(sheets, sheets.renameSheet("Budget", "Budget"));
      assertThrows(NullPointerException.class, () -> sheets.renameSheet(null, "Summary"));
      assertThrows(IllegalArgumentException.class, () -> sheets.renameSheet(" ", "Summary"));
      assertThrows(NullPointerException.class, () -> sheets.renameSheet("Budget", null));
      assertThrows(IllegalArgumentException.class, () -> sheets.renameSheet("Budget", " "));
      assertThrows(SheetNotFoundException.class, () -> sheets.renameSheet("Missing", "Summary"));
      assertThrows(IllegalArgumentException.class, () -> sheets.renameSheet("Budget", "Archive"));
      assertThrows(IllegalArgumentException.class, () -> sheets.renameSheet("Budget", "Bad/Name"));
      assertThrows(NullPointerException.class, () -> sheets.deleteSheet(null));
      assertThrows(IllegalArgumentException.class, () -> sheets.deleteSheet(" "));
      assertThrows(SheetNotFoundException.class, () -> sheets.deleteSheet("Missing"));
      sheets.deleteSheet("Archive"); // leaves only Budget; next delete must be rejected
      IllegalArgumentException lastSheet =
          assertThrows(IllegalArgumentException.class, () -> sheets.deleteSheet("Budget"));
      assertTrue(lastSheet.getMessage().contains("at least one sheet"));
      workbook.getOrCreateSheet("Archive"); // restore two-sheet state for moveSheet tests
      assertThrows(NullPointerException.class, () -> sheets.moveSheet(null, 0));
      assertThrows(IllegalArgumentException.class, () -> sheets.moveSheet(" ", 0));
      assertThrows(SheetNotFoundException.class, () -> sheets.moveSheet("Missing", 0));
      IllegalArgumentException negativeIndex =
          assertThrows(IllegalArgumentException.class, () -> sheets.moveSheet("Budget", -1));
      assertTrue(negativeIndex.getMessage().contains("workbook has"));
      IllegalArgumentException tooLargeIndex =
          assertThrows(IllegalArgumentException.class, () -> sheets.moveSheet("Budget", 2));
      assertTrue(tooLargeIndex.getMessage().contains("valid positions are 0 to 1"));
      assertThrows(
          NullPointerException.class,
          () ->
              workbook
                  .persistence()
                  .save(
                      null,
                      dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                      ExcelTempFileFactoryTestSupport.tempFileFactory()));
    }
  }

  @Test
  void rejectsNonXlsxWorkbookFiles() throws Exception {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-legacy-", ".xls");

    try (HSSFWorkbook workbook = new HSSFWorkbook();
        var outputStream = Files.newOutputStream(workbookPath)) {
      workbook.createSheet("Budget").createRow(0).createCell(0).setCellValue("Legacy");
      workbook.write(outputStream);
    }

    WorkbookNotOpenableException exception =
        assertThrows(
            WorkbookNotOpenableException.class,
            () ->
                ExcelWorkbooks.open(
                    workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory()));
    assertTrue(exception.getMessage().contains("not openable"));
  }

  @SuppressWarnings({"PMD.CloseResource", "PMD.UseTryWithResources"})
  @Test
  void validatesFormulaEnvironmentOverloadsAndCloseFailureAggregation() throws Exception {
    ExcelFormulaEnvironment defaults = ExcelFormulaEnvironment.defaults();

    assertThrows(
        NullPointerException.class,
        () ->
            ExcelWorkbooks.open(null, defaults, ExcelTempFileFactoryTestSupport.tempFileFactory()));

    Path missingPath =
        ExcelTempFiles.createManagedTempDirectory("gridgrind-missing-env-")
            .resolve("missing-with-env.xlsx");
    WorkbookNotFoundException missingWorkbook =
        assertThrows(
            WorkbookNotFoundException.class,
            () ->
                ExcelWorkbooks.open(
                    missingPath, defaults, ExcelTempFileFactoryTestSupport.tempFileFactory()));
    assertEquals(missingPath.toAbsolutePath(), missingWorkbook.workbookPath());

    Path legacyWorkbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-legacy-env-", ".xls");
    try (HSSFWorkbook workbook = new HSSFWorkbook();
        var outputStream = Files.newOutputStream(legacyWorkbookPath)) {
      workbook.createSheet("Budget").createRow(0).createCell(0).setCellValue("Legacy");
      workbook.write(outputStream);
    }
    WorkbookNotOpenableException unsupportedFormat =
        assertThrows(
            WorkbookNotOpenableException.class,
            () ->
                ExcelWorkbooks.open(
                    legacyWorkbookPath,
                    defaults,
                    ExcelTempFileFactoryTestSupport.tempFileFactory()));
    assertTrue(unsupportedFormat.getMessage().contains("not openable"));

    try (ExcelWorkbook workbook = ExcelWorkbooks.create(null)) {
      assertEquals(
          ExcelFormulaEnvironment.defaults().runtimeContext(), workbook.formulaRuntimeContext());
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Header"));
      assertThrows(
          CellNotFoundException.class,
          () -> workbook.formulas().evaluate(List.of(new ExcelFormulaCellTarget("Budget", "B1"))));
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
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("\"hello\""));
      workbook.formulas().evaluateAll();

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

      workbook.formulas().clearCaches();

      assertFalse(ctCell.isSetV());
      assertFalse(ctCell.isSetIs());
      assertFalse(ctCell.isSetT());
    }
  }

  @Test
  void clearFormulaCachesLeavesUnsetTypeMetadataUnset() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new ResolvedFormula("1+1"));
      workbook.formulas().evaluateAll();

      org.apache.poi.xssf.usermodel.XSSFCell cell =
          workbook.xssfWorkbook().getSheet("Budget").getRow(0).getCell(0);
      var ctCell = cell.getCTCell();
      if (ctCell.isSetT()) {
        ctCell.unsetT();
      }
      assertFalse(ctCell.isSetT());

      workbook.formulas().clearCaches();

      assertFalse(ctCell.isSetT());
    }
  }

  @Test
  void saveWithEmptyPersistenceOptionsActsLikePlainSave() throws IOException {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-empty-persist-", ".xlsx");
    try {
      try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
        workbook.getOrCreateSheet("Alpha").cells().setCell("A1", new CellValue.Text("Hello"));
        workbook
            .persistence()
            .save(
                workbookPath,
                dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                ExcelOoxmlPersistenceOptions.none(),
                ExcelTempFileFactoryTestSupport.tempFileFactory());
      }
      try (ExcelWorkbook reopened =
          ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
        assertEquals("Hello", reopened.sheet("Alpha").cells().text("A1"));
      }
    } finally {
      Files.deleteIfExists(workbookPath);
    }
  }

  @Test
  void repeatedEquivalentWorkbookSavesProduceIdenticalBytes() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-deterministic-save-", ".xlsx");
    try {
      byte[] firstSaveBytes = saveBudgetWorkbook(workbookPath);
      byte[] secondSaveBytes = saveBudgetWorkbook(workbookPath);

      assertArrayEquals(firstSaveBytes, secondSaveBytes);
    } finally {
      Files.deleteIfExists(workbookPath);
    }
  }

  @Test
  void createNewSaveRejectsExistingWorkbookTarget() throws IOException {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-create-new-", ".xlsx");
    try {
      saveBudgetWorkbook(workbookPath);
      try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
        workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Again"));
        assertThrows(
            FileAlreadyExistsException.class,
            () ->
                workbook
                    .persistence()
                    .save(
                        workbookPath,
                        dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.CREATE_NEW,
                        ExcelTempFileFactoryTestSupport.tempFileFactory()));
      }
    } finally {
      Files.deleteIfExists(workbookPath);
    }
  }

  @Test
  void savesToPathsWithoutParentDirectories() throws IOException {
    Path relativePath = Path.of("gridgrind-relative-" + UUID.randomUUID() + ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Hello"));
      workbook
          .persistence()
          .save(
              relativePath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
      assertTrue(Files.exists(relativePath));
    } finally {
      Files.deleteIfExists(relativePath);
    }
  }

  private static byte[] saveBudgetWorkbook(Path workbookPath) throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Hello"));
      workbook.getOrCreateSheet("Budget").cells().setCell("B1", new CellValue.NumberValue(12.0d));
      workbook.formulas().markRecalculateOnOpen();
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    return Files.readAllBytes(workbookPath);
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
    assertEquals(
        new dev.erst.gridgrind.core.model.layout.PaneReport.Frozen(1, 2, 3, 4),
        XlsxRoundTrip.pane(workbookPath, "Budget"));
  }

  @Test
  void savesAndReopensFormattingDepthStyles() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-style-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Item"));
      workbook
          .sheet("Budget")
          .cells()
          .applyStyle(
              "A1",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.of(
                      new CellAlignmentPatch(
                          Optional.of(true),
                          Optional.of(ExcelHorizontalAlignment.CENTER),
                          Optional.of(ExcelVerticalAlignment.TOP),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.of(
                      new CellFontPatch(
                          Optional.of(true),
                          Optional.of(false),
                          Optional.of("Aptos"),
                          Optional.of(new FontHeightSpec.Points(new BigDecimal("11.5"))),
                          Optional.of(ColorSpec.rgb(new RgbHex("#1F4E78"))),
                          Optional.of(true),
                          Optional.of(true))),
                  Optional.of(
                      CellFillSpec.patternForeground(
                          ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#FFF2CC")))),
                  Optional.of(
                      new CellBorderPatch(
                          Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)),
                          Optional.empty(),
                          Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.DOUBLE)),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty()));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    CellStyleReport style = XlsxRoundTrip.cellStyle(workbookPath, "Budget", "A1");
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
    assertEquals(rgb("#FFF2CC"), fillForegroundColor(style.fill()));
    assertEquals(ExcelBorderStyle.THIN, borderStyle(style.border().top()));
    assertEquals(ExcelBorderStyle.DOUBLE, borderStyle(style.border().right()));
    assertEquals(ExcelBorderStyle.THIN, borderStyle(style.border().bottom()));
    assertEquals(ExcelBorderStyle.THIN, borderStyle(style.border().left()));
  }

  @Test
  void savesAndReopensCompatibleGradientGeometryWithProtection() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-gradient-style-roundtrip-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.cells().setCell("A1", new CellValue.Text("Linear"));
      budget
          .cells()
          .applyStyle(
              "A1",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.of(
                      CellFillSpec.gradient(
                          CellGradientFillSpec.linear(
                              Optional.of(42.5d),
                              List.of(
                                  new CellGradientStopSpec(
                                      0.0d, ColorSpec.rgb(new RgbHex("#736C00"))),
                                  new CellGradientStopSpec(1.0d, ColorSpec.theme(3)))))),
                  Optional.empty(),
                  Optional.of(new CellProtectionPatch(Optional.of(true), Optional.of(true)))));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook reopenedWorkbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      CellStyleReport linearStyle =
          reopenedWorkbook.sheet("Budget").cells().snapshotCell("A1").style().orElseThrow();

      CellGradientFillReport gradient = fillGradient(linearStyle.fill());
      assertEquals("LINEAR", gradientType(gradient));
      assertEquals(42.5d, gradientDegree(gradient));
      assertNull(gradientLeft(gradient));
      assertEquals(CellColorReport.rgb(new RgbHex("#736C00")), gradient.stops().get(0).color());
      assertEquals(CellColorReport.theme(3), gradient.stops().get(1).color());
      assertTrue(linearStyle.protection().locked());
      assertTrue(linearStyle.protection().hiddenFormula());
    }
  }

  @Test
  void persistsHyperlinksCommentsAndNamedRangesAcrossSaves() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-authoring-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.cells().setCell("A1", new CellValue.Text("Report"));
      budget.cells().setCell("B4", new CellValue.NumberValue(61.0));
      budget
          .annotations()
          .setHyperlink("A1", new HyperlinkTarget.Url("https://example.com/report"));
      budget.annotations().setComment("A1", new ResolvedComment("Review", "GridGrind", true));
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("B4"))));
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("LocalItem"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2"))));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    var hyperlink = XlsxRoundTrip.cellHyperlink(workbookPath, "Budget", "A1");
    var comment = XlsxRoundTrip.cellComment(workbookPath, "Budget", "A1");
    assertEquals(new HyperlinkTarget.Url("https://example.com/report"), hyperlink.orElseThrow());
    assertAll(
        () -> assertEquals("Review", comment.orElseThrow().text()),
        () -> assertEquals("GridGrind", comment.orElseThrow().author()),
        () -> assertTrue(comment.orElseThrow().visible()));
    List<NamedRangeReport> namedRanges = XlsxRoundTrip.namedRanges(workbookPath);
    assertEquals(2, namedRanges.size());
    assertTrue(
        namedRanges.contains(
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
                new NamedRangeScopeReport.Workbook(),
                namedRanges.stream()
                    .filter(namedRange -> "BudgetTotal".equals(namedRange.name().value()))
                    .findFirst()
                    .orElseThrow()
                    .refersToFormula(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B4")))));
    assertTrue(
        namedRanges.contains(
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalItem"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
                namedRanges.stream()
                    .filter(namedRange -> "LocalItem".equals(namedRange.name().value()))
                    .findFirst()
                    .orElseThrow()
                    .refersToFormula(),
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2")))));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(2, workbook.names().namedRangeCount());
      assertEquals(2, workbook.names().namedRanges().size());
      workbook.names().deleteNamedRange("BudgetTotal", new NamedRangeScopeReport.Workbook());
      assertEquals(1, workbook.names().namedRangeCount());
    }
  }

  @Test
  void persistsLatestHyperlinkTargetAfterRepeatedWrites() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-hyperlink-replace-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("C");
      sheet
          .annotations()
          .setHyperlink("F18", new HyperlinkTarget.Email("Report_Value@example.com"));
      sheet
          .annotations()
          .setHyperlink("F18", new HyperlinkTarget.Email("Summary.Total@example.com"));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    var hyperlink = XlsxRoundTrip.cellHyperlink(workbookPath, "C", "F18");
    assertEquals(new HyperlinkTarget.Email("Summary.Total@example.com"), hyperlink.orElseThrow());
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

      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("B4"))));
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("C1"))));

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

      assertTrue(ExcelWorkbookNames.shouldExpose(workbookScoped));
      assertFalse(ExcelWorkbookNames.shouldExpose(null, false, false));
      assertFalse(ExcelWorkbookNames.shouldExpose("HiddenBudgetTotal", false, true));
      assertFalse(ExcelWorkbookNames.shouldExpose("_xlnm.Print_Area", false, false));
      assertFalse(ExcelWorkbookNames.shouldExpose("_XLNM.PRINT_TITLES", false, false));
      assertFalse(ExcelWorkbookNames.shouldExpose("BudgetFn", true, false));

      assertTrue(workbook.scopeMatches(workbookScoped, new NamedRangeScopeReport.Workbook()));
      assertFalse(
          workbook.scopeMatches(
              workbookScoped,
              new NamedRangeScopeReport.Sheet(
                  new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"))));
      assertTrue(
          workbook.scopeMatches(
              sheetScoped,
              new NamedRangeScopeReport.Sheet(
                  new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"))));
      assertFalse(workbook.scopeMatches(sheetScoped, new NamedRangeScopeReport.Workbook()));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              workbook.scopeMatches(
                  sheetScoped,
                  new NamedRangeScopeReport.Sheet(
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel("Missing"))));

      assertEquals(
          List.of(
              new NamedRangeReport.RangeReport(
                  new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
                  new NamedRangeScopeReport.Workbook(),
                  "Budget!$C$1",
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("C1"))),
              new NamedRangeReport.FormulaReport(
                  new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetRollup"),
                  new NamedRangeScopeReport.Workbook(),
                  "SUM(Budget!$B$2:$B$3)"),
              new NamedRangeReport.RangeReport(
                  new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("LocalItem"),
                  new NamedRangeScopeReport.Sheet(
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
                  "Budget!$A$1",
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1")))),
          workbook.names().namedRanges());
    }
  }

  @Test
  void validatesNamedRangeOperations() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");

      assertThrows(NullPointerException.class, () -> workbook.names().setNamedRange(null));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              workbook
                  .names()
                  .setNamedRange(
                      new StructuredMutationAction.SetNamedRange(
                          new dev.erst.gridgrind.core.vocabulary.DefinedName("A1"),
                          new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                          NamedRangeTarget.range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("B4")))));
      assertThrows(
          NullPointerException.class,
          () -> workbook.names().deleteNamedRange(null, new NamedRangeScopeReport.Workbook()));
      assertThrows(
          NullPointerException.class, () -> workbook.names().deleteNamedRange("BudgetTotal", null));
      assertThrows(
          NamedRangeNotFoundException.class,
          () ->
              workbook
                  .names()
                  .deleteNamedRange("BudgetTotal", new NamedRangeScopeReport.Workbook()));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              workbook
                  .names()
                  .setNamedRange(
                      new StructuredMutationAction.SetNamedRange(
                          new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                          new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Missing")),
                          NamedRangeTarget.range(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Missing"),
                              new dev.erst.gridgrind.core.vocabulary.RangeRef("A1")))));
    }
  }

  @Test
  void rejectsSheetNamesExceeding31Characters() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      String exactly31 = "A".repeat(31);
      String tooLong = "A".repeat(32);

      assertDoesNotThrow(() -> workbook.getOrCreateSheet(exactly31));
      assertThrows(IllegalArgumentException.class, () -> workbook.getOrCreateSheet(tooLong));
      assertThrows(IllegalArgumentException.class, () -> workbook.sheet(tooLong));
    }
  }

  @Test
  void saveCanonicalizesAmbiguousPoiColumnOutlineDefinitions() throws IOException {
    Path workbookPath = ExcelTempFiles.createManagedTempFile("gridgrind-column-save-", ".xlsx");

    try {
      try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
        XSSFSheet sheet = workbook.xssfWorkbook().createSheet("Budget");

        sheet.groupColumn(2, 3);
        for (int repetition = 0; repetition < 6; repetition++) {
          sheet.groupColumn(2, 2);
        }
        sheet.groupColumn(1, 3);

        assertFalse(
            columnDefinitionsAreCanonical(sheet),
            "raw Apache POI grouping should leave ambiguous overlapping column definitions");

        workbook
            .persistence()
            .save(
                workbookPath,
                dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
                ExcelTempFileFactoryTestSupport.tempFileFactory());
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

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }

  private static CellColorReport rgb(String rgb) {
    return CellColorReport.rgb(new RgbHex(rgb));
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
