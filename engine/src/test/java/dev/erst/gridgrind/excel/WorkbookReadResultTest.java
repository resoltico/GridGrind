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

    WorkbookReadResult.WorkbookSummaryResult workbookSummary =
        new WorkbookReadResult.WorkbookSummaryResult(
            "workbook", new WorkbookReadResult.WorkbookSummary(1, sheetNames, 1, true));
    WorkbookReadResult.NamedRangesResult namedRangesResult =
        new WorkbookReadResult.NamedRangesResult("ranges", namedRanges);
    WorkbookReadResult.SheetSummaryResult sheetSummary =
        new WorkbookReadResult.SheetSummaryResult(
            "sheet", new WorkbookReadResult.SheetSummary("Budget", 4, 3, 2));
    WorkbookReadResult.CellsResult cellsResult =
        new WorkbookReadResult.CellsResult("cells", "Budget", cells);
    WorkbookReadResult.WindowResult windowResult =
        new WorkbookReadResult.WindowResult(
            "window", new WorkbookReadResult.Window("Budget", "A1", 1, 1, rows));
    WorkbookReadResult.MergedRegionsResult mergedRegionsResult =
        new WorkbookReadResult.MergedRegionsResult("merged", "Budget", mergedRegions);
    WorkbookReadResult.HyperlinksResult hyperlinksResult =
        new WorkbookReadResult.HyperlinksResult("hyperlinks", "Budget", hyperlinks);
    WorkbookReadResult.CommentsResult commentsResult =
        new WorkbookReadResult.CommentsResult("comments", "Budget", comments);
    WorkbookReadResult.SheetLayoutResult layoutResult =
        new WorkbookReadResult.SheetLayoutResult(
            "layout",
            new WorkbookReadResult.SheetLayout(
                "Budget",
                new WorkbookReadResult.FreezePane.Frozen(1, 1, 1, 1),
                columns,
                resultRows));
    WorkbookReadResult.FormulaSurfaceResult formulaSurfaceResult =
        new WorkbookReadResult.FormulaSurfaceResult(
            "formula",
            new WorkbookReadResult.FormulaSurface(
                1, List.of(new WorkbookReadResult.SheetFormulaSurface("Budget", 1, 1, formulas))));
    WorkbookReadResult.SheetSchemaResult sheetSchemaResult =
        new WorkbookReadResult.SheetSchemaResult(
            "schema", new WorkbookReadResult.SheetSchema("Budget", "A1", 3, 2, 2, schemaColumns));
    WorkbookReadResult.NamedRangeSurfaceResult namedRangeSurfaceResult =
        new WorkbookReadResult.NamedRangeSurfaceResult(
            "surface", new WorkbookReadResult.NamedRangeSurface(1, 0, 1, 0, namedRangeEntries));

    sheetNames.clear();
    namedRanges.clear();
    cells.clear();
    rows.clear();
    mergedRegions.clear();
    hyperlinks.clear();
    comments.clear();
    columns.clear();
    resultRows.clear();
    formulas.clear();
    schemaColumns.clear();
    namedRangeEntries.clear();

    assertEquals(List.of("Budget"), workbookSummary.workbook().sheetNames());
    assertEquals(1, namedRangesResult.namedRanges().size());
    assertEquals("Budget", sheetSummary.sheet().sheetName());
    assertEquals("A1", cellsResult.cells().getFirst().address());
    assertEquals("A1", windowResult.window().rows().getFirst().cells().getFirst().address());
    assertEquals("A1:B2", mergedRegionsResult.mergedRegions().getFirst().range());
    assertEquals(
        "https://example.com/report",
        hyperlinksResult.hyperlinks().getFirst().hyperlink().target());
    assertEquals("Review", commentsResult.comments().getFirst().comment().text());
    assertInstanceOf(
        WorkbookReadResult.FreezePane.Frozen.class, layoutResult.layout().freezePanes());
    assertEquals(1, formulaSurfaceResult.analysis().totalFormulaCellCount());
    assertEquals("STRING", sheetSchemaResult.analysis().columns().getFirst().dominantType());
    assertEquals(
        WorkbookReadResult.NamedRangeBackingKind.RANGE,
        namedRangeSurfaceResult.analysis().namedRanges().getFirst().kind());
  }

  @Test
  void validatesNestedReadResultInvariants() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.WorkbookSummary(-1, List.of("Budget"), 0, true));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.WorkbookSummary(1, List.of("Budget"), -1, true));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadResult.WorkbookSummary(1, null, 0, true));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadResult.NamedRangesResult("ranges", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadResult.NamedRangesResult(
                "ranges", List.of((ExcelNamedRangeSnapshot) null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.SheetSummary("Budget", -1, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.SheetSummary("Budget", 0, -2, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.SheetSummary("Budget", 0, 0, -2));
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
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.FreezePane.Frozen(-1, 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.FreezePane.Frozen(0, -1, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.FreezePane.Frozen(0, 0, -1, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.FreezePane.Frozen(0, 0, 0, -1));
    assertInstanceOf(
        WorkbookReadResult.FreezePane.None.class, new WorkbookReadResult.FreezePane.None());
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
}
