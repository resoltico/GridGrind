package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
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
        NullPointerException.class, () -> new WorkbookCoreResult.NamedRangesResult("ranges", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookCoreResult.NamedRangesResult(
                "ranges", List.of((ExcelNamedRangeSnapshot) null)));
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
        () -> new WorkbookSheetResult.CellsResult("cells", "Budget", null));
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
        () -> new WorkbookRuleResult.DataValidationsResult("validations", "Budget", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
                "conditionalFormattingHealth", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookRuleResult.AutofiltersResult("autofilters", "Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookRuleResult.TablesResult("tables", null));
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
        () -> new WorkbookSheetResult.PrintLayoutResult("print", "Budget", null));
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
                List.of(new WorkbookSurfaceResult.TypeCount("STRING", 1)),
                "STRING"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSurfaceResult.SchemaColumn(
                0,
                "A1",
                "Item",
                -1,
                0,
                List.of(new WorkbookSurfaceResult.TypeCount("STRING", 1)),
                "STRING"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookSurfaceResult.SchemaColumn(
                0,
                "A1",
                "Item",
                0,
                -1,
                List.of(new WorkbookSurfaceResult.TypeCount("STRING", 1)),
                "STRING"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookSurfaceResult.TypeCount("STRING", 0));
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
                    "#,##0.00")),
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
        new WorkbookDrawingResult.PivotTablesResult("pivots", List.of(pivot));
    WorkbookAnalysisResult.PivotTableHealthResult pivotTableHealth =
        new WorkbookAnalysisResult.PivotTableHealthResult("pivot-health", health);

    assertEquals("Budget Pivot", pivotTables.pivotTables().getFirst().name());
    assertEquals(1, pivotTableHealth.analysis().checkedPivotTableCount());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookDrawingResult.PivotTablesResult("pivots", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookAnalysisResult.PivotTableHealthResult("pivot-health", null));
  }

  @Test
  void reportsExcelNativeIndexDiagnosticsForReadResults() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

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
                    List.of(new WorkbookSurfaceResult.TypeCount("STRING", 1)),
                    "STRING"));
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

  private static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new ExcelCellFontSnapshot(
            false,
            false,
            "Aptos",
            ExcelFontHeight.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        ExcelCellFillSnapshot.pattern(ExcelFillPattern.NONE),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }

  private ReadResultFixture createReadResultFixture() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());
    List<String> sheetNames = new ArrayList<>(List.of("Budget"));
    List<ExcelNamedRangeSnapshot> namedRanges =
        new ArrayList<>(
            List.of(
                new ExcelNamedRangeSnapshot.RangeSnapshot(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    "Budget!$B$4",
                    new ExcelNamedRangeTarget("Budget", "B4"))));
    List<ExcelCellSnapshot> cells = new ArrayList<>(List.of(blank));
    List<WorkbookSheetResult.WindowRow> rows =
        new ArrayList<>(List.of(new WorkbookSheetResult.WindowRow(0, List.of(blank))));
    List<WorkbookSheetResult.MergedRegion> mergedRegions =
        new ArrayList<>(List.of(new WorkbookSheetResult.MergedRegion("A1:B2")));
    List<WorkbookSheetResult.CellHyperlink> hyperlinks =
        new ArrayList<>(
            List.of(
                new WorkbookSheetResult.CellHyperlink(
                    "A1", new ExcelHyperlink.Url("https://example.com/report"))));
    List<WorkbookSheetResult.CellComment> comments =
        new ArrayList<>(
            List.of(
                new WorkbookSheetResult.CellComment(
                    "A1", new ExcelCommentSnapshot("Review", "GridGrind", false, null, null))));
    List<WorkbookSheetResult.ColumnLayout> columns =
        new ArrayList<>(List.of(new WorkbookSheetResult.ColumnLayout(0, 12.5, false, 0, false)));
    List<WorkbookSheetResult.RowLayout> resultRows =
        new ArrayList<>(List.of(new WorkbookSheetResult.RowLayout(0, 18.0, false, 0, false)));
    List<ExcelDataValidationSnapshot> validations =
        new ArrayList<>(
            List.of(
                new ExcelDataValidationSnapshot.Supported(
                    List.of("A2:A5"),
                    new ExcelDataValidationDefinition(
                        new ExcelDataValidationRule.TextLength(
                            ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                        true,
                        false,
                        new ExcelDataValidationPrompt(
                            "Reason", "Use 20 characters or fewer.", true),
                        null))));
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
                    List.of(new WorkbookSurfaceResult.TypeCount("STRING", 2)),
                    "STRING")));
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
            "workbook",
            new WorkbookCoreResult.WorkbookSummary.WithSheets(
                1, sheetNames, "Budget", List.of("Budget"), 1, true)),
        new WorkbookCoreResult.NamedRangesResult("ranges", namedRanges),
        new WorkbookSheetResult.SheetSummaryResult(
            "sheet",
            new WorkbookSheetResult.SheetSummary(
                "Budget",
                ExcelSheetVisibility.VISIBLE,
                new WorkbookSheetResult.SheetProtection.Unprotected(),
                4,
                3,
                2)),
        new WorkbookSheetResult.CellsResult("cells", "Budget", cells),
        new WorkbookSheetResult.WindowResult(
            "window", new WorkbookSheetResult.Window("Budget", "A1", 1, 1, rows)),
        new WorkbookSheetResult.MergedRegionsResult("merged", "Budget", mergedRegions),
        new WorkbookSheetResult.HyperlinksResult("hyperlinks", "Budget", hyperlinks),
        new WorkbookSheetResult.CommentsResult("comments", "Budget", comments),
        new WorkbookSheetResult.SheetLayoutResult(
            "layout",
            new WorkbookSheetResult.SheetLayout(
                "Budget",
                new ExcelSheetPane.Frozen(1, 1, 1, 1),
                125,
                defaultSheetPresentationSnapshot(),
                columns,
                resultRows)),
        new WorkbookSheetResult.PrintLayoutResult("print", "Budget", defaultPrintLayoutSnapshot()),
        new WorkbookRuleResult.DataValidationsResult("validations", "Budget", validations),
        new WorkbookAnalysisResult.ConditionalFormattingHealthResult(
            "conditionalFormattingHealth",
            new WorkbookAnalysis.ConditionalFormattingHealth(
                1, new WorkbookAnalysis.AnalysisSummary(0, 0, 0, 0), List.of())),
        new WorkbookRuleResult.AutofiltersResult("autofilters", "Budget", autofilters),
        new WorkbookRuleResult.TablesResult("tables", tables),
        new WorkbookSurfaceResult.FormulaSurfaceResult(
            "formula",
            new WorkbookSurfaceResult.FormulaSurface(
                1,
                List.of(new WorkbookSurfaceResult.SheetFormulaSurface("Budget", 1, 1, formulas)))),
        new WorkbookSurfaceResult.SheetSchemaResult(
            "schema",
            new WorkbookSurfaceResult.SheetSchema("Budget", "A1", 3, 2, 2, schemaColumns)),
        new WorkbookSurfaceResult.NamedRangeSurfaceResult(
            "surface", new WorkbookSurfaceResult.NamedRangeSurface(1, 0, 1, 0, namedRangeEntries)));
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
        fixture.hyperlinksResult().hyperlinks().getFirst().hyperlink().target());
    assertEquals("Review", fixture.commentsResult().comments().getFirst().comment().text());
    assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), fixture.layoutResult().layout().pane());
    assertEquals(125, fixture.layoutResult().layout().zoomPercent());
    assertEquals(defaultPrintLayoutSnapshot(), fixture.printLayoutResult().printLayout());
    assertEquals(
        "A2:A5", fixture.dataValidationsResult().validations().getFirst().ranges().getFirst());
    assertEquals(
        "conditionalFormattingHealth", fixture.conditionalFormattingHealthResult().stepId());
    assertEquals("E1:F4", fixture.autofiltersResult().autofilters().getFirst().range());
    assertEquals("BudgetTable", fixture.tablesResult().tables().getFirst().name());
    assertEquals(1, fixture.formulaSurfaceResult().analysis().totalFormulaCellCount());
    assertEquals(
        "STRING", fixture.sheetSchemaResult().analysis().columns().getFirst().dominantType());
    assertEquals(
        WorkbookSurfaceResult.NamedRangeBackingKind.RANGE,
        fixture.namedRangeSurfaceResult().analysis().namedRanges().getFirst().kind());
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

  private static ExcelPrintLayout defaultPrintLayout() {
    return new ExcelPrintLayout(
        new ExcelPrintLayout.Area.Range("A1:B20"),
        ExcelPrintOrientation.LANDSCAPE,
        new ExcelPrintLayout.Scaling.Fit(1, 0),
        new ExcelPrintLayout.TitleRows.Band(0, 0),
        new ExcelPrintLayout.TitleColumns.Band(0, 0),
        new ExcelHeaderFooterText("Budget", "", "2026"),
        new ExcelHeaderFooterText("", "Confidential", ""));
  }

  private static ExcelPrintLayoutSnapshot defaultPrintLayoutSnapshot() {
    return new ExcelPrintLayoutSnapshot(
        defaultPrintLayout(),
        new ExcelPrintSetupSnapshot(
            new ExcelPrintMarginsSnapshot(0.0d, 0.0d, 0.0d, 0.0d, 0.0d, 0.0d),
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
        null,
        ExcelSheetOutlineSummary.defaults(),
        ExcelSheetDefaults.defaults(),
        List.of());
  }
}
