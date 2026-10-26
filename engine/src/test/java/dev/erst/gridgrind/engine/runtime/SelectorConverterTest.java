package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.*;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for selector-first executor conversion seams. */
class SelectorConverterTest {
  @Test
  void convertsWorkbookSheetNamedRangeTableAndPivotSelectors() {

    assertEquals(
        new ExcelNamedRangeSelection.All(),
        SelectorConverter.toExcelNamedRangeSelection(new NamedRangeSelector.All()));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(new ExcelNamedRangeSelector.ByName("BudgetTotal"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.ByName(new DefinedName("BudgetTotal"))));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(
                new ExcelNamedRangeSelector.ByName("BudgetTotal"),
                new ExcelNamedRangeSelector.ByName("OpsTotal"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.ByNames(
                List.of(new DefinedName("BudgetTotal"), new DefinedName("OpsTotal")))));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(new ExcelNamedRangeSelector.WorkbookScope("BudgetTotal"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.SheetScope(
                new DefinedName("LocalItem"), new SheetName("Budget"))));
    assertEquals(
        new ExcelNamedRangeSelection.Selected(
            List.of(
                new ExcelNamedRangeSelector.WorkbookScope("BudgetTotal"),
                new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget"))),
        SelectorConverter.toExcelNamedRangeSelection(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
                    new NamedRangeSelector.SheetScope(
                        new DefinedName("LocalItem"), new SheetName("Budget"))))));
  }

  @Test
  void convertsCellAndRangeSelectorsAndRejectsUnsupportedShapes() {
    assertEquals(
        new SelectorConverter.SheetLocalCellSelection(
            "Budget", new ExcelCellSelection.AllUsedCells()),
        SelectorConverter.toSheetLocalCellSelection(
            new CellSelector.AllUsedInSheet(new SheetName("Budget"))));
    assertEquals(
        new SelectorConverter.SheetLocalCellSelection(
            "Budget", new ExcelCellSelection.Selected(List.of("A1"))),
        SelectorConverter.toSheetLocalCellSelection(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1"))));
    assertEquals(
        new SelectorConverter.SheetLocalCellSelection(
            "Budget", new ExcelCellSelection.Selected(List.of("A1", "B2"))),
        SelectorConverter.toSheetLocalCellSelection(
            new CellSelector.ByAddresses(
                new SheetName("Budget"), List.of(new CellAddress("A1"), new CellAddress("B2")))));
    assertEquals(
        new SelectorConverter.SheetLocalCellAddresses("Budget", List.of("A1")),
        SelectorConverter.toSheetLocalCellAddresses(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1"))));
    assertEquals(
        new SelectorConverter.SheetLocalCellAddresses("Budget", List.of("A1", "B2")),
        SelectorConverter.toSheetLocalCellAddresses(
            new CellSelector.ByAddresses(
                new SheetName("Budget"), List.of(new CellAddress("A1"), new CellAddress("B2")))));
    assertEquals(
        new SelectorConverter.SingleCellTarget("Budget", "A1"),
        SelectorConverter.toSingleCellTarget(
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1"))));

    assertEquals(
        new SelectorConverter.SheetLocalRangeSelection("Budget", new ExcelRangeSelection.All()),
        SelectorConverter.toSheetLocalRangeSelection(
            new RangeSelector.AllOnSheet(new SheetName("Budget"))));
    assertEquals(
        new SelectorConverter.SheetLocalRangeSelection(
            "Budget", new ExcelRangeSelection.Selected(List.of("A1:B2"))),
        SelectorConverter.toSheetLocalRangeSelection(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2"))));
    assertEquals(
        new SelectorConverter.SheetLocalRangeSelection(
            "Budget", new ExcelRangeSelection.Selected(List.of("A1:B2", "D4:E5"))),
        SelectorConverter.toSheetLocalRangeSelection(
            new RangeSelector.ByRanges(
                new SheetName("Budget"), List.of(new RangeRef("A1:B2"), new RangeRef("D4:E5")))));
    assertEquals(
        new SelectorConverter.SheetLocalRangeSelection(
            "Budget", new ExcelRangeSelection.Selected(List.of("B3:D4"))),
        SelectorConverter.toSheetLocalRangeSelection(
            new RangeSelector.RectangularWindow(
                new SheetName("Budget"), new CellAddress("B3"), 2, 3)));
    assertEquals(
        new SelectorConverter.SingleRangeTarget("Budget", "A1:B2"),
        SelectorConverter.toSingleRangeTarget(
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2"))));
    assertEquals(
        new SelectorConverter.SingleRangeTarget("Budget", "B3:D4"),
        SelectorConverter.toSingleRangeTarget(
            new RangeSelector.RectangularWindow(
                new SheetName("Budget"), new CellAddress("B3"), 2, 3)));

    assertEquals(
        "selector must provide exact addresses here; ALL_USED_IN_SHEET is not supported",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    SelectorConverter.toSheetLocalCellAddresses(
                        new CellSelector.AllUsedInSheet(new SheetName("Budget"))))
            .getMessage());
  }

  @Test
  void exposesSheetNamesAndBandSpansForSheetScopedSelectors() {
    assertEquals(
        "Budget", SelectorConverter.toSheetName(new SheetSelector.ByName(new SheetName("Budget"))));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(
            new DrawingObjectSelector.AllOnSheet(new SheetName("Budget"))));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(
            new DrawingObjectSelector.ByName(new SheetName("Budget"), "OpsPicture")));

    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(new RangeSelector.AllOnSheet(new SheetName("Budget"))));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(new RowBandSelector.Span(new SheetName("Budget"), 1, 3)));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(
            new RowBandSelector.Insertion(new SheetName("Budget"), 4, 2)));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(new ColumnBandSelector.Span(new SheetName("Budget"), 2, 4)));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(
            new ColumnBandSelector.Insertion(new SheetName("Budget"), 5, 1)));
    assertEquals(
        "Budget",
        SelectorConverter.toSheetName(
            new TableSelector.ByNameOnSheet(
                new TableName("BudgetTable"), new SheetName("Budget"))));
    assertEquals(
        "Report",
        SelectorConverter.toSheetName(
            new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", new SheetName("Report"))));
    assertEquals(
        new ExcelRowSpan(1, 3),
        SelectorConverter.toExcelRowSpan(new RowBandSelector.Span(new SheetName("Budget"), 1, 3)));
    assertEquals(
        new ExcelColumnSpan(2, 4),
        SelectorConverter.toExcelColumnSpan(
            new ColumnBandSelector.Span(new SheetName("Budget"), 2, 4)));
  }
}
