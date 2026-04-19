package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.*;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for selector-first executor conversion seams. */
class SelectorConverterTest {
  @Test
  void convertsWorkbookSheetNamedRangeTableAndPivotSelectors() {
    assertEquals(
        new ExcelSheetSelection.All(),
        SelectorConverter.toExcelSheetSelection(new SheetSelector.All()));
    assertEquals(
        new ExcelSheetSelection.Selected(List.of("Budget")),
        SelectorConverter.toExcelSheetSelection(new SheetSelector.ByName("Budget")));
    assertEquals(
        new ExcelSheetSelection.Selected(List.of("Budget", "Ops")),
        SelectorConverter.toExcelSheetSelection(
            new SheetSelector.ByNames(List.of("Budget", "Ops"))));

    assertEquals(
        new ExcelNamedRangeSelection.All(),
        SelectorConverter.toExcelNamedRangeSelection(new NamedRangeSelector.All()));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(new ExcelNamedRangeSelector.ByName("BudgetTotal"))),
        SelectorConverter.toExcelNamedRangeSelection(new NamedRangeSelector.ByName("BudgetTotal")));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(
                new ExcelNamedRangeSelector.ByName("BudgetTotal"),
                new ExcelNamedRangeSelector.ByName("OpsTotal"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.ByNames(List.of("BudgetTotal", "OpsTotal"))));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(new ExcelNamedRangeSelector.WorkbookScope("BudgetTotal"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.SheetScope("LocalItem", "Budget")));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(
                new ExcelNamedRangeSelector.WorkbookScope("BudgetTotal"),
                new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                    new NamedRangeSelector.SheetScope("LocalItem", "Budget")))));

    assertEquals(
        new ExcelTableSelection.All(),
        SelectorConverter.toExcelTableSelection(new TableSelector.All()));
    assertEquals(
        new ExcelTableSelection.ByNames(List.of("BudgetTable")),
        SelectorConverter.toExcelTableSelection(new TableSelector.ByName("BudgetTable")));
    assertEquals(
        new ExcelTableSelection.ByNames(List.of("BudgetTable", "OpsTable")),
        SelectorConverter.toExcelTableSelection(
            new TableSelector.ByNames(List.of("BudgetTable", "OpsTable"))));
    assertEquals(
        new ExcelTableSelection.ByNames(List.of("BudgetTable")),
        SelectorConverter.toExcelTableSelection(
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget")));

    assertEquals(
        new ExcelPivotTableSelection.All(),
        SelectorConverter.toExcelPivotTableSelection(new PivotTableSelector.All()));
    assertEquals(
        new ExcelPivotTableSelection.ByNames(List.of("Sales Pivot 2026")),
        SelectorConverter.toExcelPivotTableSelection(
            new PivotTableSelector.ByName("Sales Pivot 2026")));
    assertEquals(
        new ExcelPivotTableSelection.ByNames(List.of("Sales Pivot 2026", "Ops Pivot")),
        SelectorConverter.toExcelPivotTableSelection(
            new PivotTableSelector.ByNames(List.of("Sales Pivot 2026", "Ops Pivot"))));
    assertEquals(
        new ExcelPivotTableSelection.ByNames(List.of("Sales Pivot 2026")),
        SelectorConverter.toExcelPivotTableSelection(
            new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", "Report")));
  }

  @Test
  void convertsCellAndRangeSelectorsAndRejectsUnsupportedShapes() {
    assertEquals(
        new SelectorConverter.SheetLocalCellSelection(
            "Budget", new ExcelCellSelection.AllUsedCells()),
        SelectorConverter.toSheetLocalCellSelection(new CellSelector.AllUsedInSheet("Budget")));
    assertEquals(
        new SelectorConverter.SheetLocalCellSelection(
            "Budget", new ExcelCellSelection.Selected(List.of("A1"))),
        SelectorConverter.toSheetLocalCellSelection(new CellSelector.ByAddress("Budget", "A1")));
    assertEquals(
        new SelectorConverter.SheetLocalCellSelection(
            "Budget", new ExcelCellSelection.Selected(List.of("A1", "B2"))),
        SelectorConverter.toSheetLocalCellSelection(
            new CellSelector.ByAddresses("Budget", List.of("A1", "B2"))));
    assertEquals(
        new SelectorConverter.SheetLocalCellAddresses("Budget", List.of("A1")),
        SelectorConverter.toSheetLocalCellAddresses(new CellSelector.ByAddress("Budget", "A1")));
    assertEquals(
        new SelectorConverter.SheetLocalCellAddresses("Budget", List.of("A1", "B2")),
        SelectorConverter.toSheetLocalCellAddresses(
            new CellSelector.ByAddresses("Budget", List.of("A1", "B2"))));
    assertEquals(
        new SelectorConverter.QualifiedCellAddresses(
            List.of(
                new ExcelFormulaCellTarget("Budget", "A1"),
                new ExcelFormulaCellTarget("Ops", "B2"))),
        SelectorConverter.toQualifiedCellAddresses(
            new CellSelector.ByQualifiedAddresses(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "A1"),
                    new CellSelector.QualifiedAddress("Ops", "B2")))));
    assertEquals(
        new SelectorConverter.SingleCellTarget("Budget", "A1"),
        SelectorConverter.toSingleCellTarget(new CellSelector.ByAddress("Budget", "A1")));

    assertEquals(
        new SelectorConverter.SheetLocalRangeSelection("Budget", new ExcelRangeSelection.All()),
        SelectorConverter.toSheetLocalRangeSelection(new RangeSelector.AllOnSheet("Budget")));
    assertEquals(
        new SelectorConverter.SheetLocalRangeSelection(
            "Budget", new ExcelRangeSelection.Selected(List.of("A1:B2"))),
        SelectorConverter.toSheetLocalRangeSelection(new RangeSelector.ByRange("Budget", "A1:B2")));
    assertEquals(
        new SelectorConverter.SheetLocalRangeSelection(
            "Budget", new ExcelRangeSelection.Selected(List.of("A1:B2", "D4:E5"))),
        SelectorConverter.toSheetLocalRangeSelection(
            new RangeSelector.ByRanges("Budget", List.of("A1:B2", "D4:E5"))));
    assertEquals(
        new SelectorConverter.SheetLocalRangeSelection(
            "Budget", new ExcelRangeSelection.Selected(List.of("B3:D4"))),
        SelectorConverter.toSheetLocalRangeSelection(
            new RangeSelector.RectangularWindow("Budget", "B3", 2, 3)));
    assertEquals(
        new SelectorConverter.SingleRangeTarget("Budget", "A1:B2"),
        SelectorConverter.toSingleRangeTarget(new RangeSelector.ByRange("Budget", "A1:B2")));
    assertEquals(
        new SelectorConverter.SingleRangeTarget("Budget", "B3:D4"),
        SelectorConverter.toSingleRangeTarget(
            new RangeSelector.RectangularWindow("Budget", "B3", 2, 3)));

    IllegalArgumentException crossSheetFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                SelectorConverter.toSheetLocalCellSelection(
                    new CellSelector.ByQualifiedAddresses(
                        List.of(new CellSelector.QualifiedAddress("Budget", "A1")))));
    assertTrue(crossSheetFailure.getMessage().contains("cross-sheet exact cell selectors"));

    assertEquals(
        "selector must provide exact addresses here; ALL_USED_IN_SHEET is not supported",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorConverter.toSheetLocalCellAddresses(
                        new CellSelector.AllUsedInSheet("Budget")))
            .getMessage());
    assertEquals(
        "selector must target one sheet with exact addresses here",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorConverter.toSheetLocalCellAddresses(
                        new CellSelector.ByQualifiedAddresses(
                            List.of(new CellSelector.QualifiedAddress("Budget", "A1")))))
            .getMessage());
  }

  @Test
  void exposesSheetNamesAndBandSpansForSheetScopedSelectors() {
    assertEquals("Budget", SelectorConverter.toSheetName(new SheetSelector.ByName("Budget")));
    assertEquals(
        "Budget", SelectorConverter.toSheetName(new DrawingObjectSelector.AllOnSheet("Budget")));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(new DrawingObjectSelector.ByName("Budget", "OpsPicture")));
    assertEquals("Budget", SelectorConverter.toSheetName(new ChartSelector.AllOnSheet("Budget")));
    assertEquals("Budget", SelectorConverter.toSheetName(new RangeSelector.AllOnSheet("Budget")));
    assertEquals("Budget", SelectorConverter.toSheetName(new RowBandSelector.Span("Budget", 1, 3)));
    assertEquals(
        "Budget", SelectorConverter.toSheetName(new RowBandSelector.Insertion("Budget", 4, 2)));
    assertEquals(
        "Budget", SelectorConverter.toSheetName(new ColumnBandSelector.Span("Budget", 2, 4)));
    assertEquals(
        "Budget", SelectorConverter.toSheetName(new ColumnBandSelector.Insertion("Budget", 5, 1)));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(new TableSelector.ByNameOnSheet("BudgetTable", "Budget")));
    assertEquals(
        "Report",
        SelectorConverter.toSheetName(
            new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", "Report")));
    assertEquals(
        new ExcelRowSpan(1, 3),
        SelectorConverter.toExcelRowSpan(new RowBandSelector.Span("Budget", 1, 3)));
    assertEquals(
        new ExcelColumnSpan(2, 4),
        SelectorConverter.toExcelColumnSpan(new ColumnBandSelector.Span("Budget", 2, 4)));
  }
}
