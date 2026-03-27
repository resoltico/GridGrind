package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookReadCommand constructor invariants and copies. */
class WorkbookReadCommandTest {
  @Test
  void createsEachReadCommandVariant() {
    WorkbookReadCommand.GetWorkbookSummary workbookSummary =
        new WorkbookReadCommand.GetWorkbookSummary("workbook");
    WorkbookReadCommand.GetNamedRanges namedRanges =
        new WorkbookReadCommand.GetNamedRanges("ranges", new ExcelNamedRangeSelection.All());
    WorkbookReadCommand.GetSheetSummary sheetSummary =
        new WorkbookReadCommand.GetSheetSummary("sheet", "Budget");
    WorkbookReadCommand.GetCells cells =
        new WorkbookReadCommand.GetCells("cells", "Budget", List.of("A1"));
    WorkbookReadCommand.GetWindow window =
        new WorkbookReadCommand.GetWindow("window", "Budget", "A1", 2, 2);
    WorkbookReadCommand.GetMergedRegions mergedRegions =
        new WorkbookReadCommand.GetMergedRegions("merged", "Budget");
    WorkbookReadCommand.GetHyperlinks hyperlinks =
        new WorkbookReadCommand.GetHyperlinks(
            "hyperlinks", "Budget", new ExcelCellSelection.AllUsedCells());
    WorkbookReadCommand.GetComments comments =
        new WorkbookReadCommand.GetComments(
            "comments", "Budget", new ExcelCellSelection.Selected(List.of("A1")));
    WorkbookReadCommand.GetSheetLayout layout =
        new WorkbookReadCommand.GetSheetLayout("layout", "Budget");
    WorkbookReadCommand.GetFormulaSurface formulaSurface =
        new WorkbookReadCommand.GetFormulaSurface(
            "formula", new ExcelSheetSelection.Selected(List.of("Budget")));
    WorkbookReadCommand.GetSheetSchema schema =
        new WorkbookReadCommand.GetSheetSchema("schema", "Budget", "A1", 3, 2);
    WorkbookReadCommand.GetNamedRangeSurface namedRangeSurface =
        new WorkbookReadCommand.GetNamedRangeSurface("surface", new ExcelNamedRangeSelection.All());
    WorkbookReadCommand.AnalyzeFormulaHealth formulaHealth =
        new WorkbookReadCommand.AnalyzeFormulaHealth(
            "formulaHealth", new ExcelSheetSelection.All());
    WorkbookReadCommand.AnalyzeHyperlinkHealth hyperlinkHealth =
        new WorkbookReadCommand.AnalyzeHyperlinkHealth(
            "hyperlinkHealth", new ExcelSheetSelection.All());
    WorkbookReadCommand.AnalyzeNamedRangeHealth namedRangeHealth =
        new WorkbookReadCommand.AnalyzeNamedRangeHealth(
            "namedRangeHealth", new ExcelNamedRangeSelection.All());
    WorkbookReadCommand.AnalyzeWorkbookFindings workbookFindings =
        new WorkbookReadCommand.AnalyzeWorkbookFindings("workbookFindings");

    assertEquals("workbook", workbookSummary.requestId());
    assertInstanceOf(ExcelNamedRangeSelection.All.class, namedRanges.selection());
    assertEquals("Budget", sheetSummary.sheetName());
    assertEquals(List.of("A1"), cells.addresses());
    assertEquals("A1", window.topLeftAddress());
    assertEquals("Budget", mergedRegions.sheetName());
    assertInstanceOf(ExcelCellSelection.AllUsedCells.class, hyperlinks.selection());
    assertInstanceOf(ExcelCellSelection.Selected.class, comments.selection());
    assertEquals("Budget", layout.sheetName());
    assertInstanceOf(ExcelSheetSelection.Selected.class, formulaSurface.selection());
    assertEquals(3, schema.rowCount());
    assertInstanceOf(ExcelNamedRangeSelection.All.class, namedRangeSurface.selection());
    assertInstanceOf(ExcelSheetSelection.All.class, formulaHealth.selection());
    assertInstanceOf(ExcelSheetSelection.All.class, hyperlinkHealth.selection());
    assertInstanceOf(ExcelNamedRangeSelection.All.class, namedRangeHealth.selection());
    assertEquals("workbookFindings", workbookFindings.requestId());
  }

  @Test
  void copiesAddressListsAndRejectsInvalidInput() {
    List<String> addresses = new ArrayList<>(List.of("A1", "B2"));

    WorkbookReadCommand.GetCells cells =
        new WorkbookReadCommand.GetCells("cells", "Budget", addresses);
    addresses.clear();

    assertEquals(List.of("A1", "B2"), cells.addresses());
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadCommand.GetWorkbookSummary(null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookReadCommand.GetWorkbookSummary(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadCommand.GetNamedRanges("ranges", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadCommand.GetSheetSummary("sheet", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadCommand.GetCells("cells", "Budget", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetCells("cells", "Budget", List.of("A1", null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadCommand.GetCells("cells", "Budget", List.of("A1", "A1")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadCommand.GetWindow("window", "Budget", "A1", 0, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadCommand.GetWindow("window", "Budget", "A1", 1, 0));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetHyperlinks("links", "Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetComments("comments", "Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetFormulaSurface("formula", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadCommand.GetSheetSchema("schema", "Budget", "A1", 0, 1));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetNamedRangeSurface("surface", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.AnalyzeFormulaHealth("formulaHealth", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.AnalyzeHyperlinkHealth("hyperlinkHealth", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.AnalyzeNamedRangeHealth("namedRangeHealth", null));
  }
}
