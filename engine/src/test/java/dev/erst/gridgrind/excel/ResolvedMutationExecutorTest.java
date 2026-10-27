package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookSampleFixtures.deleteNamedRangeStep;
import static dev.erst.gridgrind.excel.WorkbookSampleFixtures.setNamedRangeStep;
import static dev.erst.gridgrind.excel.WorkbookSampleFixtures.setPivotTableStep;
import static dev.erst.gridgrind.excel.WorkbookSampleFixtures.setTableStep;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsSpec;
import dev.erst.gridgrind.core.model.layout.SheetDisplaySpec;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummarySpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Integration tests for ResolvedMutationExecutor applying core actions to a workbook. */
class ResolvedMutationExecutorTest {
  @Test
  void appliesAllSupportedCommandTypesThroughVarargs() throws IOException {
    var workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-command-layout-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Archive")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Scratch")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Ops")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
                  new ResolvedMutationAction.SetRange(
                      List.of(
                          List.of(new CellValue.Text("Item"), new CellValue.NumberValue(49.0)),
                          List.of(new CellValue.Text("Tax"), new CellValue.NumberValue(5.0))))),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B1")),
                  new WorkbookMutationAction.MergeCells()),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B1")),
                  new CellMutationAction.ApplyStyle(
                      new CellStylePatch(
                          Optional.empty(),
                          Optional.of(
                              new CellAlignmentPatch(
                                  Optional.empty(),
                                  Optional.ofNullable(ExcelHorizontalAlignment.CENTER),
                                  Optional.ofNullable(ExcelVerticalAlignment.CENTER),
                                  Optional.empty(),
                                  Optional.empty())),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty()))),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new ResolvedMutationAction.AppendRow(
                      List.of(new CellValue.Text("Total"), new ResolvedFormula("SUM(B1:B2)")))),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A2")),
                  new CellMutationAction.ClearRange()),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                  new CellMutationAction.SetHyperlink(new HyperlinkTarget.Document("Budget!B3"))),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                  new CellMutationAction.ClearHyperlink()),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B1")),
                  new ResolvedMutationAction.SetComment(
                      new ResolvedComment("Review", "GridGrind", false))),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("B1")),
                  new CellMutationAction.ClearComment()),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("C1:C3")),
                  new ResolvedMutationAction.SetDataValidation(validationDefinition())),
              new dev.erst.gridgrind.core.plan.ResolvedStep(
                  new dev.erst.gridgrind.core.plan.RangeSelector.ByRanges(
                      new SheetName("Budget"),
                      List.of("C2").stream()
                          .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                          .toList()),
                  new StructuredMutationAction.ClearDataValidations()),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B3")),
                  new StructuredMutationAction.SetAutofilter(List.of(), Optional.empty())),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new StructuredMutationAction.ClearAutofilter()),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("A1:B3")),
                  new ResolvedMutationAction.SetRange(
                      List.of(
                          List.of(new CellValue.Text("Owner"), new CellValue.Text("Task")),
                          List.of(new CellValue.Text("Ada"), new CellValue.Text("Queue")),
                          List.of(new CellValue.Text("Lin"), new CellValue.Text("Pack"))))),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("D1:E3")),
                  new ResolvedMutationAction.SetRange(
                      List.of(
                          List.of(new CellValue.Text("Region"), new CellValue.Text("Desk")),
                          List.of(new CellValue.Text("North"), new CellValue.Text("A1")),
                          List.of(new CellValue.Text("South"), new CellValue.Text("B1"))))),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("G1:H3")),
                  new ResolvedMutationAction.SetRange(
                      List.of(
                          List.of(new CellValue.Text("Stage"), new CellValue.Text("Team")),
                          List.of(new CellValue.Text("Review"), new CellValue.Text("Docs")),
                          List.of(new CellValue.Text("Ship"), new CellValue.Text("Ops"))))),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("D1:E3")),
                  new StructuredMutationAction.SetAutofilter(List.of(), Optional.empty())),
              setTableStep(
                  new ResolvedTable(
                      new TableName("Queue"),
                      new SheetName("Ops"),
                      new RangeRef("A1:B3"),
                      false,
                      new TableStyleSpec.None())),
              setTableStep(
                  new ResolvedTable(
                      new TableName("Stages"),
                      new SheetName("Ops"),
                      new RangeRef("G1:H3"),
                      false,
                      new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false))),
              new ResolvedStep(
                  new TableSelector.ByNameOnSheet(new TableName("Stages"), new SheetName("Ops")),
                  new StructuredMutationAction.DeleteTable()),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Archive")),
                  new WorkbookMutationAction.RenameSheet(new SheetName("History"))),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("History")),
                  new WorkbookMutationAction.MoveSheet(0)),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Scratch")),
                  new WorkbookMutationAction.DeleteSheet()),
              setNamedRangeStep(
                  new StructuredMutationAction.SetNamedRange(
                      new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                      new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                      NamedRangeTarget.range(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("B3")))),
              setNamedRangeStep(
                  new StructuredMutationAction.SetNamedRange(
                      new dev.erst.gridgrind.core.vocabulary.DefinedName("HistoryCell"),
                      new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("History")),
                      NamedRangeTarget.range(
                          new dev.erst.gridgrind.core.vocabulary.SheetName("History"),
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A1")))),
              deleteNamedRangeStep("BudgetTotal", new NamedRangeScopeReport.Workbook()),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new WorkbookMutationAction.AutoSizeColumns()),
              new ResolvedStep(
                  new ColumnBandSelector.Span(new SheetName("Budget"), 0, 1),
                  new WorkbookMutationAction.SetColumnWidth(16.0)),
              new ResolvedStep(
                  new RowBandSelector.Span(new SheetName("Budget"), 0, 0),
                  new WorkbookMutationAction.SetRowHeight(28.5)),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new WorkbookMutationAction.SetSheetPane(new PaneSpec.Frozen(1, 1, 1, 1))),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new WorkbookMutationAction.SetSheetZoom(125)),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new ResolvedMutationAction.SetPrintLayout(
                      new ResolvedPrintLayout(
                          new PrintAreaSpec.Range(new RangeRef("A1:B3")),
                          ExcelPrintOrientation.LANDSCAPE,
                          new PrintScalingSpec.Fit(1, 0),
                          new PrintTitleRowsSpec.Band(0, 0),
                          new PrintTitleColumnsSpec.Band(0, 0),
                          new HeaderFooterTextReport("Budget", "", ""),
                          new HeaderFooterTextReport("", "Confidential", "")))),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B1")),
                  new WorkbookMutationAction.UnmergeCells())));
      workbook.formulas().markRecalculateOnOpen();
      assertEquals(List.of("History", "Budget", "Ops"), workbook.sheets().sheetNames());
      assertEquals("Item", workbook.sheet("Budget").cells().text("A1"));
      assertEquals(54.0, workbook.sheet("Budget").cells().number("B3"));
      assertTrue(workbook.sheet("Budget").cells().snapshotCell("A1").hyperlink().isEmpty());
      assertTrue(workbook.sheet("Budget").cells().snapshotCell("B1").comment().isEmpty());
      assertEquals(1, workbook.names().namedRangeCount());
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          workbook
              .sheet("Budget")
              .cells()
              .snapshotCell("A1")
              .style()
              .orElseThrow()
              .alignment()
              .horizontalAlignment());
      assertEquals("BLANK", workbook.sheet("Budget").cells().snapshotCell("A2").type());
      assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Scratch"));
      assertTrue(workbook.formulas().recalculateOnOpenEnabled());
      dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult autofilters =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("autofilters"),
                          new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Ops")),
                          new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                              .GetAutofilters())));
      dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult tables =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("tables"),
                          new TableSelector.All(),
                          new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                              .GetTables())));
      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("D1:E3")),
              new AutofilterEntryReport.TableOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"), "Queue")),
          autofilters.autofilters());
      assertEquals(
          List.of("Queue"), tables.tables().stream().map(table -> table.name().value()).toList());
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(List.of(), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(
        new dev.erst.gridgrind.core.model.layout.PaneReport.Frozen(1, 1, 1, 1),
        XlsxRoundTrip.pane(workbookPath, "Budget"));
    assertEquals(125, XlsxRoundTrip.zoomPercent(workbookPath, "Budget"));
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        XlsxRoundTrip.printLayout(workbookPath, "Budget").orientation());
    assertEquals(
        List.of(
            new DataValidationEntryReport.Supported(
                List.of("C1", "C3").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                    .toList(),
                dev.erst.gridgrind.excel.DataValidationFixtures.report(validationDefinition()))),
        XlsxRoundTrip.dataValidations(workbookPath, "Budget"));
    assertEquals(
        List.of(
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("HistoryCell"),
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("History")),
                "History!$A$1",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("History"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("A1")))),
        XlsxRoundTrip.namedRanges(workbookPath));
    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult autofilters =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      reopened,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("autofilters"),
                          new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Ops")),
                          new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                              .GetAutofilters())));
      dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult tables =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      reopened,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("tables"),
                          new TableSelector.All(),
                          new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                              .GetTables())));
      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("D1:E3")),
              new AutofilterEntryReport.TableOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"), "Queue")),
          autofilters.autofilters());
      assertEquals(1, tables.tables().size());
      assertEquals("Queue", tables.tables().getFirst().name().value());
      assertTrue(tables.tables().getFirst().behavior().hasAutofilter());
    }
  }

  @Test
  void rejectsMutationCommandsWhenSheetDoesNotExist() throws IOException {
    WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new CellSelector.ByAddress(new SheetName("Missing"), new CellAddress("A1")),
                      new ResolvedMutationAction.SetCell(new CellValue.Text("x")))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new RangeSelector.ByRange(new SheetName("Missing"), new RangeRef("A1")),
                      new ResolvedMutationAction.SetRange(
                          List.of(List.of(new CellValue.Text("x")))))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new CellSelector.ByAddress(new SheetName("Missing"), new CellAddress("A1")),
                      new CellMutationAction.SetHyperlink(
                          new HyperlinkTarget.Document("Budget!A1")))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new CellSelector.ByAddress(new SheetName("Missing"), new CellAddress("A1")),
                      new ResolvedMutationAction.SetComment(
                          new ResolvedComment("Review", "GridGrind", false)))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new RangeSelector.ByRange(new SheetName("Missing"), new RangeRef("A1")),
                      new CellMutationAction.ApplyStyle(
                          new CellStylePatch(
                              Optional.empty(),
                              Optional.of(
                                  new CellAlignmentPatch(
                                      Optional.empty(),
                                      Optional.ofNullable(ExcelHorizontalAlignment.CENTER),
                                      Optional.ofNullable(ExcelVerticalAlignment.TOP),
                                      Optional.empty(),
                                      Optional.empty())),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty())))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new RangeSelector.ByRange(new SheetName("Missing"), new RangeRef("A1")),
                      new ResolvedMutationAction.SetDataValidation(validationDefinition()))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new RangeSelector.ByRanges(
                          new SheetName("Missing"),
                          List.of("A1").stream()
                              .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                              .toList()),
                      new StructuredMutationAction.SetConditionalFormatting(
                          new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                              List.of(
                                  new ConditionalFormattingRuleSpec.FormulaRule(
                                      "A1>0",
                                      false,
                                      Optional.of(
                                          new DifferentialStylePatch(
                                              Optional.empty(),
                                              Optional.of(true),
                                              Optional.empty(),
                                              Optional.empty(),
                                              Optional.empty(),
                                              Optional.empty(),
                                              Optional.empty(),
                                              Optional.empty(),
                                              Optional.empty())))))))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new dev.erst.gridgrind.core.plan.ResolvedStep(
                      new dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet(
                          new SheetName("Missing")),
                      new StructuredMutationAction.ClearDataValidations())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new dev.erst.gridgrind.core.plan.ResolvedStep(
                      new dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet(
                          new SheetName("Missing")),
                      new StructuredMutationAction.ClearConditionalFormatting())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new RangeSelector.ByRange(new SheetName("Missing"), new RangeRef("A1:B2")),
                      new StructuredMutationAction.SetAutofilter(List.of(), Optional.empty()))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new SheetSelector.ByName(new SheetName("Missing")),
                      new StructuredMutationAction.ClearAutofilter())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  setTableStep(
                      new ResolvedTable(
                          new TableName("Queue"),
                          new SheetName("Missing"),
                          new RangeRef("A1:B2"),
                          false,
                          new TableStyleSpec.None()))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new TableSelector.ByNameOnSheet(
                          new TableName("Queue"), new SheetName("Missing")),
                      new StructuredMutationAction.DeleteTable())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new ResolvedStep(
                      new SheetSelector.ByName(new SheetName("Missing")),
                      new ResolvedMutationAction.AppendRow(List.of(new CellValue.Text("x"))))));
    }
  }

  @Test
  void appliesSheetPresentationThroughSheetStructureDispatch() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();
      workbook.getOrCreateSheet("Budget");
      SheetPresentationSpec presentation =
          new SheetPresentationSpec(
              new SheetDisplaySpec(false, false, false, true, true),
              Optional.of(ColorSpec.rgb(new RgbHex("#225577"))),
              new SheetOutlineSummarySpec(false, false),
              new SheetDefaultsSpec(12, 19.5d),
              List.of(
                  new IgnoredErrorSpec(
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("B2:A1"),
                      List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))));

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Budget")),
                  new WorkbookMutationAction.SetSheetPresentation(presentation))));

      SheetLayoutReport layout = workbook.sheet("Budget").layout().snapshot();
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetDisplayReport(
              false, false, false, true, true),
          layout.presentation().display());
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#225577"))),
          layout.presentation().tabColor());
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport(false, false),
          layout.presentation().outlineSummary());
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetDefaultsReport(12, 19.5d),
          layout.presentation().sheetDefaults());
      assertEquals(
          List.of(
              new dev.erst.gridgrind.core.model.validation.IgnoredErrorReport(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"),
                  List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))),
          layout.presentation().ignoredErrors());
    }
  }

  @Test
  void appliesPivotTableCommandsAndSurfacesReadback() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

      executor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Data")),
              new WorkbookMutationAction.EnsureSheet()),
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Report")),
              new WorkbookMutationAction.EnsureSheet()),
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Data"), new RangeRef("A1:D5")),
              new ResolvedMutationAction.SetRange(
                  List.of(
                      List.of(
                          new CellValue.Text("Region"),
                          new CellValue.Text("Stage"),
                          new CellValue.Text("Owner"),
                          new CellValue.Text("Amount")),
                      List.of(
                          new CellValue.Text("North"),
                          new CellValue.Text("Plan"),
                          new CellValue.Text("Ada"),
                          new CellValue.NumberValue(10)),
                      List.of(
                          new CellValue.Text("North"),
                          new CellValue.Text("Do"),
                          new CellValue.Text("Ada"),
                          new CellValue.NumberValue(15)),
                      List.of(
                          new CellValue.Text("South"),
                          new CellValue.Text("Plan"),
                          new CellValue.Text("Lin"),
                          new CellValue.NumberValue(7)),
                      List.of(
                          new CellValue.Text("South"),
                          new CellValue.Text("Do"),
                          new CellValue.Text("Lin"),
                          new CellValue.NumberValue(12))))),
          setPivotTableStep(
              new PivotTableSpec(
                  "Ops Pivot",
                  new dev.erst.gridgrind.core.vocabulary.SheetName("Report"),
                  new PivotTableSpec.PivotSourceSpec.Range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Data"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:D5")),
                  new PivotTableSpec.Anchor(
                      new dev.erst.gridgrind.core.vocabulary.CellAddress("A3")),
                  List.of("Region"),
                  List.of("Stage"),
                  List.of(),
                  List.of(
                      new PivotTableSpec.DataField(
                          "Amount",
                          ExcelPivotDataConsolidateFunction.SUM,
                          "Total Amount",
                          Optional.of("#,##0.00"))))));

      dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult pivotTables =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("pivots"),
                          new PivotTableSelector.All(),
                          new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                              .GetPivotTables())));
      PivotTableReport.Supported pivot =
          assertInstanceOf(PivotTableReport.Supported.class, pivotTables.pivotTables().getFirst());

      assertEquals("Ops Pivot", pivot.name());
      assertEquals("Amount", pivot.dataFields().getFirst().sourceColumnName());

      executor.apply(
          workbook,
          new ResolvedStep(
              new PivotTableSelector.ByNameOnSheet("Ops Pivot", new SheetName("Report")),
              new StructuredMutationAction.DeletePivotTable()));

      dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult afterDelete =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.PivotTablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("pivots"),
                          new PivotTableSelector.All(),
                          new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                              .GetPivotTables())));
      assertEquals(List.of(), afterDelete.pivotTables());
    }
  }

  @Test
  void appliesDrawingCommandsThroughTheWorkbookCommandSurface() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();
      DrawingAnchorSpec.TwoCell firstAnchor =
          new DrawingAnchorSpec.TwoCell(
              new DrawingMarkerSpec(1, 1, 0, 0),
              new DrawingMarkerSpec(4, 6, 0, 0),
              ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
      DrawingAnchorSpec.TwoCell movedAnchor =
          new DrawingAnchorSpec.TwoCell(
              new DrawingMarkerSpec(6, 2, 0, 0),
              new DrawingMarkerSpec(9, 7, 0, 0),
              ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      byte[] pngBytes =
          java.util.Base64.getDecoder()
              .decode(
                  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");

      executor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()),
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new ResolvedMutationAction.SetPicture(
                  new ResolvedPicture(
                      "OpsPicture",
                      new BinaryContent(pngBytes),
                      ExcelPictureFormat.PNG,
                      firstAnchor,
                      Optional.of("Queue preview")))),
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new ResolvedMutationAction.SetShape(
                  new ResolvedShape.SimpleShape(
                      "OpsShape", firstAnchor, "rect", Optional.of("Queue")))),
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new ResolvedMutationAction.SetEmbeddedObject(
                  new ResolvedEmbeddedObject(
                      "OpsEmbed",
                      "Payload",
                      "payload.txt",
                      "payload.txt",
                      new BinaryContent(
                          "payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                      ExcelPictureFormat.PNG,
                      new BinaryContent(pngBytes),
                      firstAnchor))),
          new ResolvedStep(
              new DrawingObjectSelector.ByName(new SheetName("Ops"), "OpsShape"),
              new DrawingMutationAction.SetDrawingObjectAnchor(movedAnchor)),
          new ResolvedStep(
              new DrawingObjectSelector.ByName(new SheetName("Ops"), "OpsPicture"),
              new DrawingMutationAction.DeleteDrawingObject()));

      List<DrawingObjectReport> drawingObjects = workbook.sheet("Ops").drawings().drawingObjects();
      assertEquals(
          List.of("OpsShape", "OpsEmbed"),
          drawingObjects.stream().map(DrawingObjectReport::name).toList());
      assertEquals(
          DrawingGeometryFixtures.report(movedAnchor),
          assertInstanceOf(DrawingObjectReport.Shape.class, drawingObjects.getFirst()).anchor());
    }
  }

  @Test
  void appliesWorkbookProtectionCommands() throws IOException {
    WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");

      executor.apply(
          workbook,
          new ResolvedStep(
              new WorkbookSelector.Current(),
              new ResolvedMutationAction.SetWorkbookProtection(
                  new ResolvedWorkbookProtection(
                      true, false, true, Optional.of("secret"), Optional.of("review")))));

      assertEquals(
          new WorkbookProtectionReport(true, false, true, true, true),
          workbook.workbookProtection());

      executor.apply(
          workbook,
          new ResolvedStep(
              new WorkbookSelector.Current(),
              new WorkbookMutationAction.ClearWorkbookProtection()));

      assertEquals(
          new WorkbookProtectionReport(false, false, false, false, false),
          workbook.workbookProtection());
    }
  }

  @Test
  void appliesClearPrintLayoutThroughIterableCommands() throws IOException {
    WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook
          .sheet("Budget")
          .layout()
          .setPrintLayout(
              new ResolvedPrintLayout(
                  new PrintAreaSpec.Range(new RangeRef("A1:B3")),
                  ExcelPrintOrientation.LANDSCAPE,
                  new PrintScalingSpec.Fit(1, 0),
                  new PrintTitleRowsSpec.Band(0, 0),
                  new PrintTitleColumnsSpec.Band(0, 0),
                  new HeaderFooterTextReport("Budget", "", ""),
                  new HeaderFooterTextReport("", "Confidential", "")));

      assertSame(
          workbook,
          executor.apply(
              workbook,
              List.of(
                  new ResolvedStep(
                      new SheetSelector.ByName(new SheetName("Budget")),
                      new WorkbookMutationAction.ClearPrintLayout()))));
      ResolvedPrintLayout clearedPrintLayout = workbook.sheet("Budget").layout().printLayout();
      assertEquals(new PrintAreaSpec.None(), clearedPrintLayout.printArea());
      assertEquals(ExcelPrintOrientation.PORTRAIT, clearedPrintLayout.orientation());
      assertEquals(new PrintScalingSpec.Automatic(), clearedPrintLayout.scaling());
      assertEquals(new PrintTitleRowsSpec.None(), clearedPrintLayout.repeatingRows());
      assertEquals(new PrintTitleColumnsSpec.None(), clearedPrintLayout.repeatingColumns());
      assertEquals(HeaderFooterTextReport.blank(), clearedPrintLayout.header());
      assertEquals(HeaderFooterTextReport.blank(), clearedPrintLayout.footer());
    }
  }

  @Test
  void appliesConditionalFormattingCommandsThroughExecutor() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();
      workbook.getOrCreateSheet("Ops");

      executor.apply(
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRanges(
                  new SheetName("Ops"),
                  List.of("A1:A3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                      .toList()),
              new StructuredMutationAction.SetConditionalFormatting(
                  new dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec(
                      List.of(
                          new ConditionalFormattingRuleSpec.FormulaRule(
                              "A1>0",
                              true,
                              Optional.of(
                                  new DifferentialStylePatch(
                                      Optional.empty(),
                                      Optional.of(true),
                                      Optional.empty(),
                                      Optional.empty(),
                                      Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
                                      Optional.empty(),
                                      Optional.empty(),
                                      Optional.empty(),
                                      Optional.empty()))))))));

      dev.erst.gridgrind.core.plan.SheetInspectionResult.ConditionalFormattingResult initial =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.ConditionalFormattingResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("cf"),
                          new dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Ops")),
                          new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                              .GetConditionalFormatting())));
      assertEquals(1, initial.conditionalFormattingBlocks().size());

      executor.apply(
          workbook,
          new dev.erst.gridgrind.core.plan.ResolvedStep(
              new dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet(new SheetName("Ops")),
              new StructuredMutationAction.ClearConditionalFormatting()));

      dev.erst.gridgrind.core.plan.SheetInspectionResult.ConditionalFormattingResult cleared =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.ConditionalFormattingResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("cf"),
                          new dev.erst.gridgrind.core.plan.RangeSelector.AllOnSheet(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Ops")),
                          new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                              .GetConditionalFormatting())));
      assertEquals(List.of(), cleared.conditionalFormattingBlocks());
    }
  }

  @Test
  void appliesRowAndColumnStructuralCommandsThroughExecutor() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-structural-commands-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Layout")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Layout"), new RangeRef("A1:F6")),
                  new ResolvedMutationAction.SetRange(
                      List.of(
                          List.of(
                              new CellValue.Text("Item"),
                              new CellValue.Text("Qty"),
                              new CellValue.Text("Status"),
                              new CellValue.Text("Note"),
                              new CellValue.Text("Owner"),
                              new CellValue.Text("Flag")),
                          List.of(
                              new CellValue.Text("Hosting"),
                              new CellValue.NumberValue(42.0),
                              new CellValue.Text("Open"),
                              new CellValue.Text("Alpha"),
                              new CellValue.Text("Ada"),
                              new CellValue.Text("Y")),
                          List.of(
                              new CellValue.Text("Support"),
                              new CellValue.NumberValue(84.0),
                              new CellValue.Text("Closed"),
                              new CellValue.Text("Beta"),
                              new CellValue.Text("Lin"),
                              new CellValue.Text("N")),
                          List.of(
                              new CellValue.Text("Ops"),
                              new CellValue.NumberValue(168.0),
                              new CellValue.Text("Open"),
                              new CellValue.Text("Gamma"),
                              new CellValue.Text("Bea"),
                              new CellValue.Text("Y")),
                          List.of(
                              new CellValue.Text("QA"),
                              new CellValue.NumberValue(21.0),
                              new CellValue.Text("Queued"),
                              new CellValue.Text("Delta"),
                              new CellValue.Text("Kai"),
                              new CellValue.Text("N")),
                          List.of(
                              new CellValue.Text("Infra"),
                              new CellValue.NumberValue(7.0),
                              new CellValue.Text("Done"),
                              new CellValue.Text("Epsilon"),
                              new CellValue.Text("Mia"),
                              new CellValue.Text("Y"))))),
              new ResolvedStep(
                  new RowBandSelector.Span(new SheetName("Layout"), 1, 3),
                  new WorkbookMutationAction.GroupRows(true)),
              new ResolvedStep(
                  new RowBandSelector.Span(new SheetName("Layout"), 5, 5),
                  new WorkbookMutationAction.SetRowVisibility(true)),
              new ResolvedStep(
                  new ColumnBandSelector.Span(new SheetName("Layout"), 1, 3),
                  new WorkbookMutationAction.GroupColumns(true)),
              new ResolvedStep(
                  new ColumnBandSelector.Span(new SheetName("Layout"), 5, 5),
                  new WorkbookMutationAction.SetColumnVisibility(true)),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Moves")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Moves"), new RangeRef("A1:D3")),
                  new ResolvedMutationAction.SetRange(
                      List.of(
                          List.of(
                              new CellValue.Text("Item"),
                              new CellValue.Text("Qty"),
                              new CellValue.Text("Status"),
                              new CellValue.Text("Note")),
                          List.of(
                              new CellValue.Text("Hosting"),
                              new CellValue.NumberValue(42.0),
                              new CellValue.Text("Open"),
                              new CellValue.Text("Alpha")),
                          List.of(
                              new CellValue.Text("Support"),
                              new CellValue.NumberValue(84.0),
                              new CellValue.Text("Closed"),
                              new CellValue.Text("Beta"))))),
              new ResolvedStep(
                  new RowBandSelector.Insertion(new SheetName("Moves"), 1, 1),
                  new WorkbookMutationAction.InsertRows()),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Moves"), new CellAddress("A2")),
                  new ResolvedMutationAction.SetCell(new CellValue.Text("Spacer"))),
              new ResolvedStep(
                  new RowBandSelector.Span(new SheetName("Moves"), 2, 3),
                  new WorkbookMutationAction.ShiftRows(1)),
              new ResolvedStep(
                  new RowBandSelector.Span(new SheetName("Moves"), 2, 2),
                  new WorkbookMutationAction.DeleteRows()),
              new ResolvedStep(
                  new ColumnBandSelector.Insertion(new SheetName("Moves"), 1, 1),
                  new WorkbookMutationAction.InsertColumns()),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Moves"), new CellAddress("B1")),
                  new ResolvedMutationAction.SetCell(new CellValue.Text("Pad"))),
              new ResolvedStep(
                  new ColumnBandSelector.Span(new SheetName("Moves"), 2, 4),
                  new WorkbookMutationAction.ShiftColumns(1)),
              new ResolvedStep(
                  new ColumnBandSelector.Span(new SheetName("Moves"), 2, 2),
                  new WorkbookMutationAction.DeleteColumns())));

      assertEquals("Pad", workbook.sheet("Moves").cells().text("B1"));
      assertEquals("Hosting", workbook.sheet("Moves").cells().text("A3"));
      assertEquals(42.0, workbook.sheet("Moves").cells().number("C3"));
      assertEquals("Beta", workbook.sheet("Moves").cells().text("E4"));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    SheetLayoutReport layout = XlsxRoundTrip.sheetLayout(workbookPath, "Layout");
    assertEquals(6, layout.rows().size());
    assertTrue(layout.rows().get(1).hidden());
    assertEquals(1, layout.rows().get(1).outlineLevel());
    assertTrue(layout.rows().get(4).collapsed());
    assertTrue(layout.rows().get(5).hidden());
    assertEquals(6, layout.columns().size());
    assertTrue(layout.columns().get(1).hidden());
    assertEquals(1, layout.columns().get(1).outlineLevel());
    assertTrue(layout.columns().get(4).collapsed());
    assertTrue(layout.columns().get(5).hidden());

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Pad", reopened.sheet("Moves").cells().text("B1"));
      assertEquals("Hosting", reopened.sheet("Moves").cells().text("A3"));
      assertEquals(42.0, reopened.sheet("Moves").cells().number("C3"));
      assertEquals("Beta", reopened.sheet("Moves").cells().text("E4"));
    }
  }

  @Test
  void appliesSheetManagementCommandsThroughExecutor() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Alpha")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Beta")),
                  new WorkbookMutationAction.EnsureSheet()),
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Alpha"), new CellAddress("A1")),
                  new ResolvedMutationAction.SetCell(new CellValue.Text("Live"))),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Alpha")),
                  new WorkbookMutationAction.CopySheet(
                      new SheetName("Alpha Copy"), new SheetCopyPosition.AtIndex(1))),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Alpha")),
                  new WorkbookMutationAction.SetActiveSheet()),
              new ResolvedStep(
                  new SheetSelector.ByNames(
                      List.of("Alpha", "Beta").stream().map(SheetName::new).toList()),
                  new WorkbookMutationAction.SetSelectedSheets()),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Beta")),
                  new WorkbookMutationAction.SetSheetVisibility(ExcelSheetVisibility.HIDDEN)),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Alpha")),
                  new ResolvedMutationAction.SetSheetProtection(
                      protectionSettings(), Optional.empty())),
              new ResolvedStep(
                  new SheetSelector.ByName(new SheetName("Alpha")),
                  new WorkbookMutationAction.ClearSheetProtection())));

      assertEquals(List.of("Alpha", "Alpha Copy", "Beta"), workbook.sheets().sheetNames());
      assertEquals("Live", workbook.sheet("Alpha Copy").cells().text("A1"));

      WorkbookSummary.WithSheets summary =
          assertInstanceOf(WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals(
          new dev.erst.gridgrind.core.vocabulary.SheetLabel("Alpha"), summary.activeSheetName());
      assertEquals(
          List.of("Alpha").stream()
              .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::new)
              .toList(),
          summary.selectedSheetNames());
      assertEquals(ExcelSheetVisibility.HIDDEN, workbook.sheetSummary("Beta").visibility());
      assertInstanceOf(
          SheetProtectionReport.Unprotected.class, workbook.sheetSummary("Alpha").protection());
    }
  }

  @Test
  void appliesUngroupCommandsThroughExecutor() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-command-ungroup-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

      executor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Layout")),
              new WorkbookMutationAction.EnsureSheet()),
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Layout"), new RangeRef("A1:D4")),
              new ResolvedMutationAction.SetRange(
                  List.of(
                      List.of(
                          new CellValue.Text("Item"),
                          new CellValue.Text("Qty"),
                          new CellValue.Text("Status"),
                          new CellValue.Text("Owner")),
                      List.of(
                          new CellValue.Text("Hosting"),
                          new CellValue.NumberValue(42.0),
                          new CellValue.Text("Open"),
                          new CellValue.Text("Ada")),
                      List.of(
                          new CellValue.Text("Support"),
                          new CellValue.NumberValue(84.0),
                          new CellValue.Text("Closed"),
                          new CellValue.Text("Lin")),
                      List.of(
                          new CellValue.Text("Ops"),
                          new CellValue.NumberValue(7.0),
                          new CellValue.Text("Done"),
                          new CellValue.Text("Mia"))))),
          new ResolvedStep(
              new RowBandSelector.Span(new SheetName("Layout"), 1, 3),
              new WorkbookMutationAction.GroupRows(true)),
          new ResolvedStep(
              new ColumnBandSelector.Span(new SheetName("Layout"), 1, 3),
              new WorkbookMutationAction.GroupColumns(true)),
          new ResolvedStep(
              new RowBandSelector.Span(new SheetName("Layout"), 1, 3),
              new WorkbookMutationAction.UngroupRows()),
          new ResolvedStep(
              new ColumnBandSelector.Span(new SheetName("Layout"), 1, 3),
              new WorkbookMutationAction.UngroupColumns()));

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    SheetLayoutReport layout = XlsxRoundTrip.sheetLayout(workbookPath, "Layout");

    assertTrue(layout.rows().size() >= 4);
    assertFalse(layout.rows().get(1).hidden());
    assertEquals(0, layout.rows().get(1).outlineLevel());
    assertTrue(layout.columns().size() >= 4);
    assertFalse(layout.columns().get(1).hidden());
    assertEquals(0, layout.columns().get(1).outlineLevel());
  }

  @Test
  void routesDedicatedHelperFamiliesAcrossWorkbookSeams() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Budget")),
              new WorkbookMutationAction.EnsureSheet()));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      assertEquals(List.of("Budget", "Ops"), workbook.sheets().sheetNames());

      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
              new ResolvedMutationAction.SetRange(
                  List.of(
                      List.of(
                          new CellValue.Text("Extremely long workflow heading"),
                          new CellValue.Text("State")),
                      List.of(new CellValue.Text("Ada"), new CellValue.Text("Queued"))))));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new ColumnBandSelector.Span(new SheetName("Budget"), 0, 0),
              new WorkbookMutationAction.SetColumnWidth(4.0)));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Budget")),
              new WorkbookMutationAction.SetSheetZoom(135)));

      double widthBeforeAutoSize =
          workbook.sheet("Budget").layout().snapshot().columns().getFirst().widthCharacters();
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Budget")),
              new WorkbookMutationAction.AutoSizeColumns()));

      assertEquals(135, workbook.sheet("Budget").layout().snapshot().zoomPercent());
      assertTrue(
          workbook.sheet("Budget").layout().snapshot().columns().getFirst().widthCharacters()
              > widthBeforeAutoSize);

      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
              new ResolvedMutationAction.SetComment(
                  new ResolvedComment("Review", "GridGrind", false))));
      ResolvedMutationExecutor.apply(
          workbook,
          setNamedRangeStep(
              new StructuredMutationAction.SetNamedRange(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetStatus"),
                  new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                  NamedRangeTarget.range(
                      new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("B2")))));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Budget")),
              new ResolvedMutationAction.SetShape(
                  new ResolvedShape.SimpleShape(
                      "BudgetShape",
                      new DrawingAnchorSpec.TwoCell(
                          new DrawingMarkerSpec(0, 0, 0, 0),
                          new DrawingMarkerSpec(2, 3, 0, 0),
                          ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                      "rect",
                      Optional.of("Queue")))));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1")),
              new CellMutationAction.ApplyStyle(
                  new CellStylePatch(
                      Optional.empty(),
                      Optional.of(
                          new CellAlignmentPatch(
                              Optional.empty(),
                              Optional.ofNullable(ExcelHorizontalAlignment.CENTER),
                              Optional.ofNullable(ExcelVerticalAlignment.CENTER),
                              Optional.empty(),
                              Optional.empty())),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty()))));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
              new StructuredMutationAction.SetAutofilter(List.of(), Optional.empty())));

      assertAll(
          () ->
              assertEquals(
                  "Review",
                  workbook
                      .sheet("Budget")
                      .cells()
                      .snapshotCell("A1")
                      .comment()
                      .orElseThrow()
                      .text()),
          () ->
              assertEquals(
                  "GridGrind",
                  workbook
                      .sheet("Budget")
                      .cells()
                      .snapshotCell("A1")
                      .comment()
                      .orElseThrow()
                      .author()),
          () ->
              assertFalse(
                  workbook
                      .sheet("Budget")
                      .cells()
                      .snapshotCell("A1")
                      .comment()
                      .orElseThrow()
                      .visible()));
      assertEquals(1, workbook.names().namedRangeCount());
      assertEquals(1, workbook.sheet("Budget").drawings().drawingObjects().size());
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          workbook
              .sheet("Budget")
              .cells()
              .snapshotCell("A1")
              .style()
              .orElseThrow()
              .alignment()
              .horizontalAlignment());
      dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult autofilters =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("autofilters"),
                          new dev.erst.gridgrind.core.plan.SheetSelector.ByName(
                              new dev.erst.gridgrind.core.vocabulary.SheetName("Budget")),
                          new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery
                              .GetAutofilters())));
      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"))),
          autofilters.autofilters());
    }
  }

  @Test
  void validatesNullWorkbooksCommandsAndCommandEntries() throws IOException {
    WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      assertThrows(
          NullPointerException.class, () -> executor.apply(workbook, (ResolvedStep[]) null));
      assertThrows(NullPointerException.class, () -> executor.apply(null, List.of()));
      assertThrows(
          NullPointerException.class,
          () -> executor.apply(workbook, (Iterable<ResolvedStep>) null));
      assertThrows(
          NullPointerException.class,
          () -> executor.apply(workbook, Arrays.asList((ResolvedStep) null)));
    }
  }

  private static ResolvedDataValidation validationDefinition() {
    return new ResolvedDataValidation(
        new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
        true,
        false,
        Optional.of(new ResolvedValidationPrompt("Status", "Pick one workflow state.", true)),
        Optional.of(
            new ResolvedValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP,
                "Invalid status",
                "Use one of the allowed values.",
                true)));
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }
}
