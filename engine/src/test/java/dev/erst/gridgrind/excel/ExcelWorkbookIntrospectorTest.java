package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for ExcelWorkbookIntrospector workbook-fact reads and named-range selection. */
class ExcelWorkbookIntrospectorTest {
  @Test
  void executesEveryIntrospectionCommandAgainstWorkbookState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.setCell("A1", ExcelCellValue.text("Report"));
      budget.setCell("B4", ExcelCellValue.number(61.0));
      budget.mergeCells("A1:B1");
      budget.setColumnWidth(0, 0, 12.5);
      budget.setRowHeight(0, 0, 18.0);
      budget.freezePanes(1, 1, 1, 1);
      budget.setHyperlink("A1", new ExcelHyperlink.Url("https://example.com/report"));
      budget.setComment("A1", new ExcelComment("Review", "GridGrind", false));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "BudgetTotal",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Budget", "B4")));

      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();

      WorkbookReadResult.WorkbookSummaryResult workbookSummary =
          cast(
              WorkbookReadResult.WorkbookSummaryResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook")));
      WorkbookReadResult.NamedRangesResult namedRanges =
          cast(
              WorkbookReadResult.NamedRangesResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetNamedRanges(
                      "ranges", new ExcelNamedRangeSelection.All())));
      WorkbookReadResult.SheetSummaryResult sheetSummary =
          cast(
              WorkbookReadResult.SheetSummaryResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetSheetSummary("sheet", "Budget")));
      WorkbookReadResult.CellsResult cells =
          cast(
              WorkbookReadResult.CellsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetCells("cells", "Budget", List.of("A1", "B4"))));
      WorkbookReadResult.WindowResult window =
          cast(
              WorkbookReadResult.WindowResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetWindow("window", "Budget", "A1", 2, 2)));
      WorkbookReadResult.MergedRegionsResult mergedRegions =
          cast(
              WorkbookReadResult.MergedRegionsResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetMergedRegions("merged", "Budget")));
      WorkbookReadResult.HyperlinksResult hyperlinks =
          cast(
              WorkbookReadResult.HyperlinksResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetHyperlinks(
                      "hyperlinks", "Budget", new ExcelCellSelection.Selected(List.of("A1")))));
      WorkbookReadResult.CommentsResult comments =
          cast(
              WorkbookReadResult.CommentsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetComments(
                      "comments", "Budget", new ExcelCellSelection.AllUsedCells())));
      WorkbookReadResult.SheetLayoutResult layout =
          cast(
              WorkbookReadResult.SheetLayoutResult.class,
              introspector.execute(
                  workbook, new WorkbookReadCommand.GetSheetLayout("layout", "Budget")));

      assertEquals(List.of("Budget"), workbookSummary.workbook().sheetNames());
      assertEquals("BudgetTotal", namedRanges.namedRanges().getFirst().name());
      assertEquals("Budget", sheetSummary.sheet().sheetName());
      assertEquals("A1", cells.cells().getFirst().address());
      assertEquals("A1", window.window().rows().getFirst().cells().getFirst().address());
      assertEquals("A1:B1", mergedRegions.mergedRegions().getFirst().range());
      assertEquals(
          "https://example.com/report", hyperlinks.hyperlinks().getFirst().hyperlink().target());
      assertEquals("Review", comments.comments().getFirst().comment().text());
      assertInstanceOf(WorkbookReadResult.FreezePane.Frozen.class, layout.layout().freezePanes());
    }
  }

  @Test
  void selectsNamedRangesByExactSelectorsAndRejectsMissingSelectors() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-introspector-ranges-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");
      workbook.createSheet("Forecast");
      var workbookName = workbook.createName();
      workbookName.setNameName("BudgetRollup");
      workbookName.setRefersToFormula("SUM(Budget!$B$2:$B$3)");
      var sheetName = workbook.createName();
      sheetName.setNameName("LocalItem");
      sheetName.setSheetIndex(0);
      sheetName.setRefersToFormula("Budget!$A$1");
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();

      List<ExcelNamedRangeSnapshot> all =
          introspector.selectNamedRanges(workbook, new ExcelNamedRangeSelection.All());
      List<ExcelNamedRangeSnapshot> selected =
          introspector.selectNamedRanges(
              workbook,
              new ExcelNamedRangeSelection.Selected(
                  List.of(
                      new ExcelNamedRangeSelector.ByName("budgetrollup"),
                      new ExcelNamedRangeSelector.WorkbookScope("BudgetRollup"),
                      new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget"))));
      assertEquals(2, all.size());
      assertEquals(List.of(all.getFirst(), all.get(1)), selected);
      NamedRangeNotFoundException missing =
          assertThrows(
              NamedRangeNotFoundException.class,
              () ->
                  introspector.selectNamedRanges(
                      workbook,
                      new ExcelNamedRangeSelection.Selected(
                          List.of(new ExcelNamedRangeSelector.ByName("MissingRange")))));
      assertEquals("MissingRange", missing.name());

      NamedRangeNotFoundException missingWorkbookScope =
          assertThrows(
              NamedRangeNotFoundException.class,
              () ->
                  introspector.selectNamedRanges(
                      workbook,
                      new ExcelNamedRangeSelection.Selected(
                          List.of(new ExcelNamedRangeSelector.WorkbookScope("MissingWorkbook")))));
      assertEquals("MissingWorkbook", missingWorkbookScope.name());
      assertInstanceOf(ExcelNamedRangeScope.WorkbookScope.class, missingWorkbookScope.scope());

      NamedRangeNotFoundException missingSheetScope =
          assertThrows(
              NamedRangeNotFoundException.class,
              () ->
                  introspector.selectNamedRanges(
                      workbook,
                      new ExcelNamedRangeSelection.Selected(
                          List.of(
                              new ExcelNamedRangeSelector.SheetScope("SharedName", "Budget")))));
      assertEquals("SharedName", missingSheetScope.name());
      ExcelNamedRangeScope.SheetScope scope =
          assertInstanceOf(ExcelNamedRangeScope.SheetScope.class, missingSheetScope.scope());
      assertEquals("Budget", scope.sheetName());
    }
  }

  @Test
  void matchSelectorHandlesSameNameScopeEdgeCases() {
    List<ExcelNamedRangeSnapshot> namedRanges =
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "SharedName",
                new ExcelNamedRangeScope.WorkbookScope(),
                "Budget!$A$1",
                new ExcelNamedRangeTarget("Budget", "A1")),
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "SharedName",
                new ExcelNamedRangeScope.SheetScope("Forecast"),
                "Forecast!$A$1",
                new ExcelNamedRangeTarget("Forecast", "A1")));

    List<ExcelNamedRangeSnapshot> workbookScoped =
        ExcelWorkbookIntrospector.matchSelector(
            namedRanges, new ExcelNamedRangeSelector.WorkbookScope("SharedName"));
    List<ExcelNamedRangeSnapshot> sheetScoped =
        ExcelWorkbookIntrospector.matchSelector(
            namedRanges, new ExcelNamedRangeSelector.SheetScope("SharedName", "Forecast"));

    assertEquals(1, workbookScoped.size());
    assertInstanceOf(ExcelNamedRangeScope.WorkbookScope.class, workbookScoped.getFirst().scope());
    assertEquals(1, sheetScoped.size());
    ExcelNamedRangeScope.SheetScope scope =
        assertInstanceOf(ExcelNamedRangeScope.SheetScope.class, sheetScoped.getFirst().scope());
    assertEquals("Forecast", scope.sheetName());

    NamedRangeNotFoundException missingWorkbookSelector =
        assertThrows(
            NamedRangeNotFoundException.class,
            () ->
                ExcelWorkbookIntrospector.matchSelector(
                    namedRanges, new ExcelNamedRangeSelector.WorkbookScope("Missing")));
    assertEquals("Missing", missingWorkbookSelector.name());

    NamedRangeNotFoundException wrongSheetSelector =
        assertThrows(
            NamedRangeNotFoundException.class,
            () ->
                ExcelWorkbookIntrospector.matchSelector(
                    namedRanges, new ExcelNamedRangeSelector.SheetScope("SharedName", "Budget")));
    assertEquals("SharedName", wrongSheetSelector.name());
  }

  @Test
  void rejectsNullWorkbookAndCommands() {
    ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();

    assertThrows(
        NullPointerException.class,
        () -> introspector.execute(null, new WorkbookReadCommand.GetWorkbookSummary("workbook")));
    assertThrows(
        NullPointerException.class, () -> introspector.execute(ExcelWorkbook.create(), null));
    assertThrows(
        NullPointerException.class,
        () -> introspector.selectNamedRanges(null, new ExcelNamedRangeSelection.All()));
    assertThrows(
        NullPointerException.class,
        () -> introspector.selectNamedRanges(ExcelWorkbook.create(), null));
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }
}
