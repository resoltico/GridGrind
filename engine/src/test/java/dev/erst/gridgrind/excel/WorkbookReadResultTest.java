package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

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
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                -1, List.of("Budget"), "Budget", List.of("Budget"), 0, true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget"), -1, true));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, null, "Budget", List.of("Budget"), 0, true));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadResult.NamedRangesResult("ranges", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadResult.NamedRangesResult(
                "ranges", List.of((ExcelNamedRangeSnapshot) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.SheetSummary(
                "Budget",
                ExcelSheetVisibility.VISIBLE,
                new WorkbookReadResult.SheetProtection.Unprotected(),
                -1,
                0,
                0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.SheetSummary(
                "Budget",
                ExcelSheetVisibility.VISIBLE,
                new WorkbookReadResult.SheetProtection.Unprotected(),
                0,
                -2,
                0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.SheetSummary(
                "Budget",
                ExcelSheetVisibility.VISIBLE,
                new WorkbookReadResult.SheetProtection.Unprotected(),
                0,
                0,
                -2));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.CellsResult("cells", "Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.Window("Budget", "A1", 0, 1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.Window("Budget", "A1", 1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadResult.WindowRow(-1, List.of(blank)));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookReadResult.MergedRegion(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadResult.CellHyperlink("A1", null));
    assertThrows(NullPointerException.class, () -> new WorkbookReadResult.CellComment("A1", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.DataValidationsResult("validations", "Budget", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadResult.ConditionalFormattingHealthResult(
                "conditionalFormattingHealth", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.AutofiltersResult("autofilters", "Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadResult.TablesResult("tables", null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(0, -1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelSheetPane.Frozen(0, 0, 0, -1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.SheetLayout(
                "Budget", new ExcelSheetPane.None(), 9, List.of(), List.of()));
    assertInstanceOf(ExcelSheetPane.None.class, new ExcelSheetPane.None());
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.PrintLayoutResult("print", "Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadResult.ColumnLayout(-1, 1.0));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookReadResult.ColumnLayout(0, 0.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.ColumnLayout(0, Double.POSITIVE_INFINITY));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookReadResult.RowLayout(-1, 1.0));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookReadResult.RowLayout(0, 0.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.RowLayout(0, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadResult.FormulaSurface(-1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.SheetFormulaSurface("Budget", -1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.SheetFormulaSurface("Budget", 0, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.FormulaPattern("SUM(B2:B3)", 0, List.of("B4")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.FormulaPattern("SUM(B2:B3)", 1, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.SheetSchema("Budget", "A1", 0, 1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.SheetSchema("Budget", "A1", 1, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.SheetSchema("Budget", "A1", 1, 1, -1, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.SchemaColumn(
                -1,
                "A1",
                "Item",
                0,
                0,
                List.of(new WorkbookReadResult.TypeCount("STRING", 1)),
                "STRING"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.SchemaColumn(
                0,
                "A1",
                "Item",
                -1,
                0,
                List.of(new WorkbookReadResult.TypeCount("STRING", 1)),
                "STRING"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.SchemaColumn(
                0,
                "A1",
                "Item",
                0,
                -1,
                List.of(new WorkbookReadResult.TypeCount("STRING", 1)),
                "STRING"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadResult.TypeCount("STRING", 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.NamedRangeSurface(-1, 0, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.NamedRangeSurface(0, -1, 0, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.NamedRangeSurface(0, 0, -1, 0, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.NamedRangeSurface(0, 0, 0, -1, List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.NamedRangeSurface(0, 0, 0, 0, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadResult.NamedRangeSurfaceEntry(
                "BudgetTotal",
                null,
                "Budget!$B$4",
                WorkbookReadResult.NamedRangeBackingKind.RANGE));
  }

  @Test
  void validatesWorkbookSummaryStateSpecificInvariants() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.WorkbookSummary.Empty(1, List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                2, List.of("Budget"), "Budget", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                0, List.of(), "Budget", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget"), -1, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Missing", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of(), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Missing"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                2, List.of("Budget", "Budget"), "Budget", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, List.of(" "), " ", List.of("Budget"), 0, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget", "Budget"), 0, false));
  }

  private static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        false,
        false,
        false,
        ExcelHorizontalAlignment.GENERAL,
        ExcelVerticalAlignment.BOTTOM,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        null,
        false,
        false,
        null,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE);
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
    List<WorkbookReadResult.WindowRow> rows =
        new ArrayList<>(List.of(new WorkbookReadResult.WindowRow(0, List.of(blank))));
    List<WorkbookReadResult.MergedRegion> mergedRegions =
        new ArrayList<>(List.of(new WorkbookReadResult.MergedRegion("A1:B2")));
    List<WorkbookReadResult.CellHyperlink> hyperlinks =
        new ArrayList<>(
            List.of(
                new WorkbookReadResult.CellHyperlink(
                    "A1", new ExcelHyperlink.Url("https://example.com/report"))));
    List<WorkbookReadResult.CellComment> comments =
        new ArrayList<>(
            List.of(
                new WorkbookReadResult.CellComment(
                    "A1", new ExcelComment("Review", "GridGrind", false))));
    List<WorkbookReadResult.ColumnLayout> columns =
        new ArrayList<>(List.of(new WorkbookReadResult.ColumnLayout(0, 12.5)));
    List<WorkbookReadResult.RowLayout> resultRows =
        new ArrayList<>(List.of(new WorkbookReadResult.RowLayout(0, 18.0)));
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
    List<WorkbookReadResult.FormulaPattern> formulas =
        new ArrayList<>(
            List.of(new WorkbookReadResult.FormulaPattern("SUM(B2:B3)", 1, List.of("B4"))));
    List<WorkbookReadResult.SchemaColumn> schemaColumns =
        new ArrayList<>(
            List.of(
                new WorkbookReadResult.SchemaColumn(
                    0,
                    "A1",
                    "Item",
                    2,
                    0,
                    List.of(new WorkbookReadResult.TypeCount("STRING", 2)),
                    "STRING")));
    List<WorkbookReadResult.NamedRangeSurfaceEntry> namedRangeEntries =
        new ArrayList<>(
            List.of(
                new WorkbookReadResult.NamedRangeSurfaceEntry(
                    "BudgetTotal",
                    new ExcelNamedRangeScope.WorkbookScope(),
                    "Budget!$B$4",
                    WorkbookReadResult.NamedRangeBackingKind.RANGE)));

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
        new WorkbookReadResult.WorkbookSummaryResult(
            "workbook",
            new WorkbookReadResult.WorkbookSummary.WithSheets(
                1, sheetNames, "Budget", List.of("Budget"), 1, true)),
        new WorkbookReadResult.NamedRangesResult("ranges", namedRanges),
        new WorkbookReadResult.SheetSummaryResult(
            "sheet",
            new WorkbookReadResult.SheetSummary(
                "Budget",
                ExcelSheetVisibility.VISIBLE,
                new WorkbookReadResult.SheetProtection.Unprotected(),
                4,
                3,
                2)),
        new WorkbookReadResult.CellsResult("cells", "Budget", cells),
        new WorkbookReadResult.WindowResult(
            "window", new WorkbookReadResult.Window("Budget", "A1", 1, 1, rows)),
        new WorkbookReadResult.MergedRegionsResult("merged", "Budget", mergedRegions),
        new WorkbookReadResult.HyperlinksResult("hyperlinks", "Budget", hyperlinks),
        new WorkbookReadResult.CommentsResult("comments", "Budget", comments),
        new WorkbookReadResult.SheetLayoutResult(
            "layout",
            new WorkbookReadResult.SheetLayout(
                "Budget", new ExcelSheetPane.Frozen(1, 1, 1, 1), 125, columns, resultRows)),
        new WorkbookReadResult.PrintLayoutResult("print", "Budget", defaultPrintLayout()),
        new WorkbookReadResult.DataValidationsResult("validations", "Budget", validations),
        new WorkbookReadResult.ConditionalFormattingHealthResult(
            "conditionalFormattingHealth",
            new WorkbookAnalysis.ConditionalFormattingHealth(
                1, new WorkbookAnalysis.AnalysisSummary(0, 0, 0, 0), List.of())),
        new WorkbookReadResult.AutofiltersResult("autofilters", "Budget", autofilters),
        new WorkbookReadResult.TablesResult("tables", tables),
        new WorkbookReadResult.FormulaSurfaceResult(
            "formula",
            new WorkbookReadResult.FormulaSurface(
                1, List.of(new WorkbookReadResult.SheetFormulaSurface("Budget", 1, 1, formulas)))),
        new WorkbookReadResult.SheetSchemaResult(
            "schema", new WorkbookReadResult.SheetSchema("Budget", "A1", 3, 2, 2, schemaColumns)),
        new WorkbookReadResult.NamedRangeSurfaceResult(
            "surface", new WorkbookReadResult.NamedRangeSurface(1, 0, 1, 0, namedRangeEntries)));
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
    assertEquals(defaultPrintLayout(), fixture.printLayoutResult().printLayout());
    assertEquals(
        "A2:A5", fixture.dataValidationsResult().validations().getFirst().ranges().getFirst());
    assertEquals(
        "conditionalFormattingHealth", fixture.conditionalFormattingHealthResult().requestId());
    assertEquals("E1:F4", fixture.autofiltersResult().autofilters().getFirst().range());
    assertEquals("BudgetTable", fixture.tablesResult().tables().getFirst().name());
    assertEquals(1, fixture.formulaSurfaceResult().analysis().totalFormulaCellCount());
    assertEquals(
        "STRING", fixture.sheetSchemaResult().analysis().columns().getFirst().dominantType());
    assertEquals(
        WorkbookReadResult.NamedRangeBackingKind.RANGE,
        fixture.namedRangeSurfaceResult().analysis().namedRanges().getFirst().kind());
  }

  private record ReadResultFixture(
      List<String> sheetNames,
      List<ExcelNamedRangeSnapshot> namedRanges,
      List<ExcelCellSnapshot> cells,
      List<WorkbookReadResult.WindowRow> rows,
      List<WorkbookReadResult.MergedRegion> mergedRegions,
      List<WorkbookReadResult.CellHyperlink> hyperlinks,
      List<WorkbookReadResult.CellComment> comments,
      List<WorkbookReadResult.ColumnLayout> columns,
      List<WorkbookReadResult.RowLayout> resultRows,
      List<ExcelDataValidationSnapshot> validations,
      List<ExcelAutofilterSnapshot> autofilters,
      List<ExcelTableSnapshot> tables,
      List<WorkbookReadResult.FormulaPattern> formulas,
      List<WorkbookReadResult.SchemaColumn> schemaColumns,
      List<WorkbookReadResult.NamedRangeSurfaceEntry> namedRangeEntries,
      WorkbookReadResult.WorkbookSummaryResult workbookSummary,
      WorkbookReadResult.NamedRangesResult namedRangesResult,
      WorkbookReadResult.SheetSummaryResult sheetSummary,
      WorkbookReadResult.CellsResult cellsResult,
      WorkbookReadResult.WindowResult windowResult,
      WorkbookReadResult.MergedRegionsResult mergedRegionsResult,
      WorkbookReadResult.HyperlinksResult hyperlinksResult,
      WorkbookReadResult.CommentsResult commentsResult,
      WorkbookReadResult.SheetLayoutResult layoutResult,
      WorkbookReadResult.PrintLayoutResult printLayoutResult,
      WorkbookReadResult.DataValidationsResult dataValidationsResult,
      WorkbookReadResult.ConditionalFormattingHealthResult conditionalFormattingHealthResult,
      WorkbookReadResult.AutofiltersResult autofiltersResult,
      WorkbookReadResult.TablesResult tablesResult,
      WorkbookReadResult.FormulaSurfaceResult formulaSurfaceResult,
      WorkbookReadResult.SheetSchemaResult sheetSchemaResult,
      WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurfaceResult) {
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
}
