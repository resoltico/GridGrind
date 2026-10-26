package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.print.*;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintMarginsReport;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSnapshot;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationSnapshot;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookReadResult constructor invariants and defensive copies. */
class WorkbookReadResultTest {
  @Test
  void createsEachResultVariantAndCopiesCollections() {
    ReadResultFixture fixture = createReadResultFixture();
    fixture.clearSources();
    assertCopiedFixture(fixture);
  }

  @Test
  void validatesNestedReadResultInvariants() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                -1, List.of("Budget"), "Budget", List.of("Budget"), 0, true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget"), -1, true));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, null, "Budget", List.of("Budget"), 0, true));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCoreResult.NamedRangesResult(new StepId("ranges"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCoreResult.NamedRangesResult(
                new StepId("ranges"), List.of((ExcelNamedRangeSnapshot) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSheetResult.SheetSummary(
                "Budget",
                ExcelSheetVisibility.VISIBLE,
                new WorkbookSheetResult.SheetProtection.Unprotected(),
                -1,
                0,
                0));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookSheetResult.CellsResult(
                new StepId("cells"), "Budget", null, WorkbookReadTestSupport.projection(), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.Window("Budget", "A1", 0, 1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.Window("Budget", "A1", 1, 0, List.of()));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookSheetResult.MergedRegion(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookSheetResult.CellHyperlink("A1", null));
    assertThrows(NullPointerException.class, () -> new WorkbookSheetResult.CellComment("A1", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookRuleResult.DataValidationsResult(
                new StepId("validations"), "Budget", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
                new StepId("conditionalFormattingHealth"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookRuleResult.AutofiltersResult(new StepId("autofilters"), "Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookRuleResult.TablesResult(new StepId("tables"), null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(0, -1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(0, 0, 0, -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSheetResult.SheetLayout(
                "Budget",
                new ExcelSheetPane.None(),
                9,
                defaultSheetPresentationSnapshot(),
                List.of(),
                List.of()));
    assertInstanceOf(ExcelSheetPane.None.class, new ExcelSheetPane.None());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSheetResult.PrintLayoutResult(new StepId("print"), "Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.ColumnLayout(-1, 1.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.ColumnLayout(0, 0.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.ColumnLayout(0, Double.POSITIVE_INFINITY, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.ColumnLayout(0, 1.0, false, -1, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.RowLayout(-1, 1.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.RowLayout(0, 0.0, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.RowLayout(0, Double.POSITIVE_INFINITY, false, 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSheetResult.RowLayout(0, 1.0, false, -1, false));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.FormulaSurface(-1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetFormulaSurface("Budget", -1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetFormulaSurface("Budget", 0, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.FormulaPattern("SUM(B2:B3)", 0, List.of("B4")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSurfaceResult.FormulaPattern("SUM(B2:B3)", 1, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 0, 1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 1, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 1, 1, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSurfaceResult.SchemaColumn(
                -1,
                "A1",
                "Item",
                0,
                0,
                List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 1)),
                "TEXT"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSurfaceResult.SchemaColumn(
                0,
                "A1",
                "Item",
                -1,
                0,
                List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 1)),
                "TEXT"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSurfaceResult.SchemaColumn(
                0,
                "A1",
                "Item",
                0,
                -1,
                List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 1)),
                "TEXT"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookSurfaceResult.TypeCount("TEXT", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(-1, 0, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(0, -1, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(0, 0, -1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(0, 0, 0, -1, List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookSurfaceResult.NamedRangeSurface(0, 0, 0, 0, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookSurfaceResult.NamedRangeSurfaceEntry(
                "BudgetTotal",
                null,
                "Budget!$B$4",
                WorkbookSurfaceResult.NamedRangeBackingKind.RANGE));
  }

  @Test
  void pivotReadResultsCopyEntriesAndRejectInvalidState() {
    ExcelPivotTableSnapshot.Supported pivot =
        new ExcelPivotTableSnapshot.Supported(
            "Budget Pivot",
            "Report",
            new ExcelPivotTableSnapshot.Anchor("A3", "A3:C8"),
            new ExcelPivotTableSnapshot.Source.Range("Data", "A1:D5"),
            List.of(new ExcelPivotTableSnapshot.Field(0, "Region")),
            List.of(),
            List.of(),
            List.of(
                new ExcelPivotTableSnapshot.DataField(
                    3,
                    "Amount",
                    ExcelPivotDataConsolidateFunction.SUM,
                    "Total Amount",
                    Optional.of("#,##0.00"))),
            false);
    WorkbookAnalysis.AnalysisFinding finding =
        new WorkbookAnalysis.AnalysisFinding(
            AnalysisFindingCode.PIVOT_TABLE_MISSING_NAME,
            AnalysisSeverity.WARNING,
            "Pivot table name is missing",
            "GridGrind assigned a synthetic identifier for readback.",
            new WorkbookAnalysis.AnalysisLocation.Sheet("Report"),
            List.of("_GG_PIVOT_Report_A3"));
    WorkbookAnalysis.PivotTableHealth health =
        new WorkbookAnalysis.PivotTableHealth(
            1, new WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0), List.of(finding));

    WorkbookDrawingResult.PivotTablesResult pivotTables =
        new WorkbookDrawingResult.PivotTablesResult(new StepId("pivots"), List.of(pivot));
    WorkbookAnalysisResult.PivotTableHealthResult pivotTableHealth =
        new WorkbookAnalysisResult.PivotTableHealthResult(new StepId("pivot-health"), health);

    assertEquals("Budget Pivot", pivotTables.pivotTables().getFirst().name());
    assertEquals(1, pivotTableHealth.analysis().checkedPivotTableCount());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingResult.PivotTablesResult(new StepId("pivots"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysisResult.PivotTableHealthResult(new StepId("pivot-health"), null));
  }

  @Test
  void reportsExcelNativeIndexDiagnosticsForReadResults() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

    IllegalArgumentException windowRowFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookSheetResult.WindowRow(-1, List.of(blank)));
    assertTrue(windowRowFailure.getMessage().contains("rowIndex -1"));
    assertTrue(windowRowFailure.getMessage().contains("Excel row 1"));

    IllegalArgumentException columnLayoutFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookSheetResult.ColumnLayout(-1, 1.0, false, 0, false));
    assertTrue(columnLayoutFailure.getMessage().contains("columnIndex -1"));
    assertTrue(columnLayoutFailure.getMessage().contains("Excel column A"));

    IllegalArgumentException rowLayoutFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> new WorkbookSheetResult.RowLayout(-1, 1.0, false, 0, false));
    assertTrue(rowLayoutFailure.getMessage().contains("rowIndex -1"));
    assertTrue(rowLayoutFailure.getMessage().contains("Excel row 1"));

    IllegalArgumentException schemaColumnFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new WorkbookSurfaceResult.SchemaColumn(
                    -1,
                    "A1",
                    "Item",
                    0,
                    0,
                    List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 1)),
                    "TEXT"));
    assertTrue(schemaColumnFailure.getMessage().contains("columnIndex -1"));
    assertTrue(schemaColumnFailure.getMessage().contains("Excel column A"));

    IllegalArgumentException lastRowFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new WorkbookSheetResult.SheetSummary(
                    "Budget",
                    ExcelSheetVisibility.VISIBLE,
                    new WorkbookSheetResult.SheetProtection.Unprotected(),
                    0,
                    -2,
                    0));
    assertTrue(lastRowFailure.getMessage().contains("lastRowIndex -2"));
    assertTrue(lastRowFailure.getMessage().contains("empty sheets report -1"));

    IllegalArgumentException lastColumnFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new WorkbookSheetResult.SheetSummary(
                    "Budget",
                    ExcelSheetVisibility.VISIBLE,
                    new WorkbookSheetResult.SheetProtection.Unprotected(),
                    0,
                    0,
                    -2));
    assertTrue(lastColumnFailure.getMessage().contains("lastColumnIndex -2"));
    assertTrue(lastColumnFailure.getMessage().contains("empty sheets report -1"));
  }

  @Test
  void validatesWorkbookSummaryStateSpecificInvariants() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCoreResult.WorkbookSummary.Empty(1, List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                2, List.of("Budget"), "Budget", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                0, List.of(), "Budget", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget"), -1, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Missing", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of(), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Missing"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                2, List.of("Budget", "Budget"), "Budget", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, List.of(" "), " ", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget", "Budget"), 0, false));
  }

  private static CellStyleReport defaultStyle() {
    return new CellStyleReport(
        "",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            FontHeightReport.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None()),
        new CellProtectionReport(true, false));
  }

  private ReadResultFixture createReadResultFixture() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());
    List<String> sheetNames = new ArrayList<>(List.of("Budget"));
    List<ExcelNamedRangeSnapshot> namedRanges =
        new ArrayList<>(
            List.of(
                new ExcelNamedRangeSnapshot.RangeSnapshot(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    "Budget!$B$4",
                    ExcelNamedRangeTarget.range("Budget", "B4"))));
    List<ExcelCellSnapshot> cells = new ArrayList<>(List.of(blank));
    List<WorkbookSheetResult.WindowRow> rows =
        new ArrayList<>(List.of(new WorkbookSheetResult.WindowRow(0, List.of(blank))));
    List<WorkbookSheetResult.MergedRegion> mergedRegions =
        new ArrayList<>(List.of(new WorkbookSheetResult.MergedRegion("A1:B2")));
    List<WorkbookSheetResult.CellHyperlink> hyperlinks =
        new ArrayList<>(
            List.of(
                new WorkbookSheetResult.CellHyperlink(
                    "A1", new HyperlinkTarget.Url("https://example.com/report"))));
    List<WorkbookSheetResult.CellComment> comments =
        new ArrayList<>(
            List.of(
                new WorkbookSheetResult.CellComment(
                    "A1",
                    new CommentReport(
                        "Review", "GridGrind", false, Optional.empty(), Optional.empty()))));
    List<WorkbookSheetResult.ColumnLayout> columns =
        new ArrayList<>(List.of(new WorkbookSheetResult.ColumnLayout(0, 12.5, false, 0, false)));
    List<WorkbookSheetResult.RowLayout> resultRows =
        new ArrayList<>(List.of(new WorkbookSheetResult.RowLayout(0, 18.0, false, 0, false)));
    List<ExcelDataValidationSnapshot> validations =
        new ArrayList<>(
            List.of(
                new ExcelDataValidationSnapshot.Supported(
                    List.of("A2:A5"),
                    new ResolvedDataValidation(
                        new DataValidationRuleSpec.TextLength(
                            ExcelComparisonOperator.LESS_OR_EQUAL,
                            ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
                            ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                                Optional.empty(), "formula2")),
                        true,
                        false,
                        Optional.of(
                            new ResolvedValidationPrompt(
                                "Reason", "Use 20 characters or fewer.", true)),
                        Optional.empty()))));
    List<ExcelAutofilterSnapshot> autofilters =
        new ArrayList<>(
            List.of(
                new ExcelAutofilterSnapshot.SheetOwned("E1:F4"),
                new ExcelAutofilterSnapshot.TableOwned("A1:C4", "BudgetTable")));
    List<ExcelTableSnapshot> tables =
        new ArrayList<>(
            List.of(
                new ExcelTableSnapshot(
                    "BudgetTable",
                    "Budget",
                    "A1:C4",
                    1,
                    1,
                    List.of("Item", "Amount", "Billable"),
                    new ExcelTableStyleSnapshot.Named(
                        "TableStyleMedium2", false, false, true, false),
                    true)));
    List<WorkbookSurfaceResult.FormulaPattern> formulas =
        new ArrayList<>(
            List.of(new WorkbookSurfaceResult.FormulaPattern("SUM(B2:B3)", 1, List.of("B4"))));
    List<WorkbookSurfaceResult.SchemaColumn> schemaColumns =
        new ArrayList<>(
            List.of(
                new WorkbookSurfaceResult.SchemaColumn(
                    0,
                    "A1",
                    "Item",
                    2,
                    0,
                    List.of(new WorkbookSurfaceResult.TypeCount("TEXT", 2)),
                    "TEXT")));
    List<WorkbookSurfaceResult.NamedRangeSurfaceEntry> namedRangeEntries =
        new ArrayList<>(
            List.of(
                new WorkbookSurfaceResult.NamedRangeSurfaceEntry(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    "Budget!$B$4",
                    WorkbookSurfaceResult.NamedRangeBackingKind.RANGE)));

    return new ReadResultFixture(
        sheetNames,
        namedRanges,
        cells,
        rows,
        mergedRegions,
        hyperlinks,
        comments,
        columns,
        resultRows,
        validations,
        autofilters,
        tables,
        formulas,
        schemaColumns,
        namedRangeEntries,
        new WorkbookCoreResult.WorkbookSummaryResult(
            new StepId("workbook"),
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, sheetNames, "Budget", List.of("Budget"), 1, true)),
        new WorkbookCoreResult.NamedRangesResult(new StepId("ranges"), namedRanges),
        new WorkbookSheetResult.SheetSummaryResult(
            new StepId("sheet"),
            new WorkbookSheetResult.SheetSummary(
                "Budget",
                ExcelSheetVisibility.VISIBLE,
                new WorkbookSheetResult.SheetProtection.Unprotected(),
                4,
                3,
                2)),
        WorkbookReadTestSupport.cellsResult("cells", "Budget", cells),
        WorkbookReadTestSupport.windowResult(
            "window", new WorkbookSheetResult.Window("Budget", "A1", 1, 1, rows)),
        new WorkbookSheetResult.MergedRegionsResult(new StepId("merged"), "Budget", mergedRegions),
        new WorkbookSheetResult.HyperlinksResult(new StepId("hyperlinks"), "Budget", hyperlinks),
        new WorkbookSheetResult.CommentsResult(new StepId("comments"), "Budget", comments),
        new WorkbookSheetResult.SheetLayoutResult(
            new StepId("layout"),
            new WorkbookSheetResult.SheetLayout(
                "Budget",
                new ExcelSheetPane.Frozen(1, 1, 1, 1),
                125,
                defaultSheetPresentationSnapshot(),
                columns,
                resultRows)),
        new WorkbookSheetResult.PrintLayoutResult(
            new StepId("print"), "Budget", defaultPrintLayoutSnapshot()),
        new WorkbookRuleResult.DataValidationsResult(
            new StepId("validations"), "Budget", validations),
        new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
            new StepId("conditionalFormattingHealth"),
            new WorkbookAnalysis.ConditionalFormattingHealth(
                1, new WorkbookAnalysis.AnalysisSummary(0, 0, 0, 0), List.of())),
        new WorkbookRuleResult.AutofiltersResult(new StepId("autofilters"), "Budget", autofilters),
        new WorkbookRuleResult.TablesResult(new StepId("tables"), tables),
        new WorkbookSurfaceResult.FormulaSurfaceResult(
            new StepId("formula"),
            new WorkbookSurfaceResult.FormulaSurface(
                1,
                List.of(new WorkbookSurfaceResult.SheetFormulaSurface("Budget", 1, 1, formulas)))),
        WorkbookReadTestSupport.sheetSchemaResult(
            "schema",
            new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 3, 2, 2, schemaColumns)),
        new WorkbookSurfaceResult.NamedRangeSurfaceResult(
            new StepId("surface"),
            new WorkbookSurfaceResult.NamedRangeSurface(1, 0, 1, 0, namedRangeEntries)));
  }

  private void assertCopiedFixture(ReadResultFixture fixture) {
    assertEquals(List.of("Budget"), fixture.workbookSummary().workbook().sheetNames());
    assertEquals(1, fixture.namedRangesResult().namedRanges().size());
    assertEquals("Budget", fixture.sheetSummary().sheet().sheetName());
    assertEquals("A1", fixture.cellsResult().cells().getFirst().address());
    assertEquals(
        "A1", fixture.windowResult().window().rows().getFirst().cells().getFirst().address());
    assertEquals("A1:B2", fixture.mergedRegionsResult().mergedRegions().getFirst().range());
    assertEquals(
        "https://example.com/report",
        assertInstanceOf(
                HyperlinkTarget.Url.class,
                fixture.hyperlinksResult().hyperlinks().getFirst().hyperlink())
            .target());
    assertEquals("Review", fixture.commentsResult().comments().getFirst().comment().text());
    assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), fixture.layoutResult().layout().pane());
    assertEquals(125, fixture.layoutResult().layout().zoomPercent());
    assertEquals(defaultPrintLayoutSnapshot(), fixture.printLayoutResult().printLayout());
    assertEquals(
        "A2:A5", fixture.dataValidationsResult().validations().getFirst().ranges().getFirst());
    assertEquals(
        "conditionalFormattingHealth",
        fixture.conditionalFormattingHealthResult().stepId().value());
    assertEquals("E1:F4", fixture.autofiltersResult().autofilters().getFirst().range());
    assertEquals("BudgetTable", fixture.tablesResult().tables().getFirst().name());
    assertEquals(1, fixture.formulaSurfaceResult().surface().totalFormulaCellCount());
    assertEquals("TEXT", fixture.sheetSchemaResult().surface().columns().getFirst().dominantType());
    assertEquals(
        WorkbookSurfaceResult.NamedRangeBackingKind.RANGE,
        fixture.namedRangeSurfaceResult().surface().namedRanges().getFirst().kind());
  }

  private record ReadResultFixture(
      List<String> sheetNames,
      List<ExcelNamedRangeSnapshot> namedRanges,
      List<ExcelCellSnapshot> cells,
      List<WorkbookSheetResult.WindowRow> rows,
      List<WorkbookSheetResult.MergedRegion> mergedRegions,
      List<WorkbookSheetResult.CellHyperlink> hyperlinks,
      List<WorkbookSheetResult.CellComment> comments,
      List<WorkbookSheetResult.ColumnLayout> columns,
      List<WorkbookSheetResult.RowLayout> resultRows,
      List<ExcelDataValidationSnapshot> validations,
      List<ExcelAutofilterSnapshot> autofilters,
      List<ExcelTableSnapshot> tables,
      List<WorkbookSurfaceResult.FormulaPattern> formulas,
      List<WorkbookSurfaceResult.SchemaColumn> schemaColumns,
      List<WorkbookSurfaceResult.NamedRangeSurfaceEntry> namedRangeEntries,
      WorkbookCoreResult.WorkbookSummaryResult workbookSummary,
      WorkbookCoreResult.NamedRangesResult namedRangesResult,
      WorkbookSheetResult.SheetSummaryResult sheetSummary,
      WorkbookSheetResult.CellsResult cellsResult,
      WorkbookSheetResult.WindowResult windowResult,
      WorkbookSheetResult.MergedRegionsResult mergedRegionsResult,
      WorkbookSheetResult.HyperlinksResult hyperlinksResult,
      WorkbookSheetResult.CommentsResult commentsResult,
      WorkbookSheetResult.SheetLayoutResult layoutResult,
      WorkbookSheetResult.PrintLayoutResult printLayoutResult,
      WorkbookRuleResult.DataValidationsResult dataValidationsResult,
      WorkbookAnalysisResult.ConditionalFormattingHealthResult conditionalFormattingHealthResult,
      WorkbookRuleResult.AutofiltersResult autofiltersResult,
      WorkbookRuleResult.TablesResult tablesResult,
      WorkbookSurfaceResult.FormulaSurfaceResult formulaSurfaceResult,
      WorkbookSurfaceResult.SheetSchemaResult sheetSchemaResult,
      WorkbookSurfaceResult.NamedRangeSurfaceResult namedRangeSurfaceResult) {
    private void clearSources() {
      sheetNames.clear();
      namedRanges.clear();
      cells.clear();
      rows.clear();
      mergedRegions.clear();
      hyperlinks.clear();
      comments.clear();
      columns.clear();
      resultRows.clear();
      validations.clear();
      autofilters.clear();
      tables.clear();
      formulas.clear();
      schemaColumns.clear();
      namedRangeEntries.clear();
    }
  }

  private static PrintLayoutReport defaultPrintLayoutSnapshot() {
    return new PrintLayoutReport(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"),
        new PrintAreaReport.Range(new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B20")),
        ExcelPrintOrientation.LANDSCAPE,
        new PrintScalingReport.Fit(1, 0),
        new PrintTitleRowsReport.Band(0, 0),
        new PrintTitleColumnsReport.Band(0, 0),
        new HeaderFooterTextReport("Budget", "", "2026"),
        new HeaderFooterTextReport("", "Confidential", ""),
        new PrintSetupReport(
            new PrintMarginsReport(0.0d, 0.0d, 0.0d, 0.0d, 0.0d, 0.0d),
            false,
            false,
            false,
            0,
            false,
            false,
            0,
            false,
            0,
            List.of(),
            List.of()));
  }

  private static ExcelSheetPresentationSnapshot defaultSheetPresentationSnapshot() {
    return new ExcelSheetPresentationSnapshot(
        ExcelSheetDisplay.defaults(),
        Optional.empty(),
        ExcelSheetOutlineSummary.defaults(),
        ExcelSheetDefaults.defaults(),
        List.of());
  }
}
