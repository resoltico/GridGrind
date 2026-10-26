package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.vocabulary.ExcelReadLimits;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookReadCommand constructor invariants and copies. */
class WorkbookReadCommandTest {
  @Test
  void createsEachReadCommandVariant() {
    WorkbookReadCommand.GetWorkbookSummary workbookSummary =
        new WorkbookReadCommand.GetWorkbookSummary(new StepId("workbook"));
    WorkbookReadCommand.GetNamedRanges namedRanges =
        new WorkbookReadCommand.GetNamedRanges(
            new StepId("ranges"), new ExcelNamedRangeSelection.All());
    WorkbookReadCommand.GetSheetSummary sheetSummary =
        new WorkbookReadCommand.GetSheetSummary(new StepId("sheet"), "Budget");
    WorkbookReadCommand.GetCells cells =
        WorkbookReadTestSupport.getCells("cells", "Budget", List.of("A1"));
    WorkbookReadCommand.GetWindow window =
        WorkbookReadTestSupport.getWindow("window", "Budget", "A1", 2, 2);
    WorkbookReadCommand.GetMergedRegions mergedRegions =
        new WorkbookReadCommand.GetMergedRegions(new StepId("merged"), "Budget");
    WorkbookReadCommand.GetHyperlinks hyperlinks =
        new WorkbookReadCommand.GetHyperlinks(
            new StepId("hyperlinks"), "Budget", new ExcelCellSelection.AllUsedCells());
    WorkbookReadCommand.GetComments comments =
        new WorkbookReadCommand.GetComments(
            new StepId("comments"), "Budget", new ExcelCellSelection.Selected(List.of("A1")));
    WorkbookReadCommand.GetSheetLayout layout =
        new WorkbookReadCommand.GetSheetLayout(new StepId("layout"), "Budget");
    WorkbookReadCommand.GetDataValidations dataValidations =
        new WorkbookReadCommand.GetDataValidations(
            new StepId("validations"), "Budget", new ExcelRangeSelection.All());
    WorkbookReadCommand.GetAutofilters autofilters =
        new WorkbookReadCommand.GetAutofilters(new StepId("autofilters"), "Budget");
    WorkbookReadCommand.GetTables tables =
        new WorkbookReadCommand.GetTables(new StepId("tables"), new TableSelector.All());
    WorkbookReadCommand.GetPivotTables pivotTables =
        new WorkbookReadCommand.GetPivotTables(new StepId("pivots"), new PivotTableSelector.All());
    WorkbookReadCommand.GetFormulaSurface formulaSurface =
        new WorkbookReadCommand.GetFormulaSurface(
            new StepId("formula"),
            new dev.erst.gridgrind.core.plan.SheetSelector.ByNames(
                List.of("Budget").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.SheetName::new)
                    .toList()));
    WorkbookReadCommand.GetSheetSchema schema =
        WorkbookReadTestSupport.getSheetSchema("schema", "Budget", "A1", 3, 2);
    WorkbookReadCommand.GetNamedRangeSurface namedRangeSurface =
        new WorkbookReadCommand.GetNamedRangeSurface(
            new StepId("surface"), new ExcelNamedRangeSelection.All());
    WorkbookReadCommand.AnalyzeFormulaHealth formulaHealth =
        new WorkbookReadCommand.AnalyzeFormulaHealth(
            new StepId("formulaHealth"), new SheetSelector.All());
    WorkbookReadCommand.AnalyzeDataValidationHealth dataValidationHealth =
        new WorkbookReadCommand.AnalyzeDataValidationHealth(
            new StepId("dataValidationHealth"), new SheetSelector.All());
    WorkbookReadCommand.AnalyzeConditionalFormattingHealth conditionalFormattingHealth =
        new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
            new StepId("conditionalFormattingHealth"), new SheetSelector.All());
    WorkbookReadCommand.AnalyzeAutofilterHealth autofilterHealth =
        new WorkbookReadCommand.AnalyzeAutofilterHealth(
            new StepId("autofilterHealth"), new SheetSelector.All());
    WorkbookReadCommand.AnalyzeTableHealth tableHealth =
        new WorkbookReadCommand.AnalyzeTableHealth(
            new StepId("tableHealth"),
            new TableSelector.ByNames(
                List.of("BudgetTable").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.TableName::new)
                    .toList()));
    WorkbookReadCommand.AnalyzePivotTableHealth pivotTableHealth =
        new WorkbookReadCommand.AnalyzePivotTableHealth(
            new StepId("pivotTableHealth"),
            new PivotTableSelector.ByNames(List.of("Budget Pivot")));
    WorkbookReadCommand.AnalyzeHyperlinkHealth hyperlinkHealth =
        new WorkbookReadCommand.AnalyzeHyperlinkHealth(
            new StepId("hyperlinkHealth"), new SheetSelector.All());
    WorkbookReadCommand.AnalyzeNamedRangeHealth namedRangeHealth =
        new WorkbookReadCommand.AnalyzeNamedRangeHealth(
            new StepId("namedRangeHealth"), new ExcelNamedRangeSelection.All());
    WorkbookReadCommand.AnalyzeWorkbookFindings workbookFindings =
        new WorkbookReadCommand.AnalyzeWorkbookFindings(new StepId("workbookFindings"));

    assertEquals("workbook", workbookSummary.stepId().value());
    assertInstanceOf(ExcelNamedRangeSelection.All.class, namedRanges.selection());
    assertEquals("Budget", sheetSummary.sheetName());
    assertEquals(List.of("A1"), cells.addresses());
    assertEquals("A1", window.window().topLeftAddress());
    assertEquals("Budget", mergedRegions.sheetName());
    assertInstanceOf(ExcelCellSelection.AllUsedCells.class, hyperlinks.selection());
    assertInstanceOf(ExcelCellSelection.Selected.class, comments.selection());
    assertEquals("Budget", layout.sheetName());
    assertEquals("Budget", dataValidations.sheetName());
    assertEquals("Budget", autofilters.sheetName());
    assertInstanceOf(TableSelector.All.class, tables.selection());
    assertInstanceOf(PivotTableSelector.All.class, pivotTables.selection());
    assertInstanceOf(SheetSelector.ByNames.class, formulaSurface.selection());
    assertEquals(3, schema.window().rowCount());
    assertInstanceOf(ExcelNamedRangeSelection.All.class, namedRangeSurface.selection());
    assertInstanceOf(SheetSelector.All.class, formulaHealth.selection());
    assertInstanceOf(SheetSelector.All.class, dataValidationHealth.selection());
    assertInstanceOf(SheetSelector.All.class, conditionalFormattingHealth.selection());
    assertInstanceOf(SheetSelector.All.class, autofilterHealth.selection());
    assertInstanceOf(TableSelector.ByNames.class, tableHealth.selection());
    assertInstanceOf(PivotTableSelector.ByNames.class, pivotTableHealth.selection());
    assertInstanceOf(SheetSelector.All.class, hyperlinkHealth.selection());
    assertInstanceOf(ExcelNamedRangeSelection.All.class, namedRangeHealth.selection());
    assertEquals("workbookFindings", workbookFindings.stepId().value());
  }

  @Test
  void copiesAddressListsAndRejectsInvalidInput() {
    List<String> addresses = new ArrayList<>(List.of("A1", "B2"));

    WorkbookReadCommand.GetCells cells =
        WorkbookReadTestSupport.getCells("cells", "Budget", addresses);
    addresses.clear();

    assertEquals(List.of("A1", "B2"), cells.addresses());
    assertThrows(
        NullPointerException.class, () -> new WorkbookReadCommand.GetWorkbookSummary(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadCommand.GetWorkbookSummary(new StepId(" ")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetNamedRanges(new StepId("ranges"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetSheetSummary(new StepId("sheet"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookReadTestSupport.getCells("cells", "Budget", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> WorkbookReadTestSupport.getCells("cells", "Budget", List.of("A1", null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookReadTestSupport.getCells("cells", "Budget", List.of("A1", "A1")));
    assertEquals(
        "addresses must not exceed 250000 but was 250001",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    WorkbookReadTestSupport.getCells(
                        "cells",
                        "Budget",
                        java.util.stream.IntStream.rangeClosed(
                                1, ExcelReadLimits.MAX_READ_CELLS + 1)
                            .mapToObj(index -> "A" + index)
                            .toList()))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookReadTestSupport.getWindow("window", "Budget", "A1", 0, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookReadTestSupport.getWindow("window", "Budget", "A1", 1, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookReadTestSupport.getWindow("window", "Budget", "A1", 1000, 1000));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookReadTestSupport.getSheetSchema("schema", "Budget", "A1", 1000, 1000));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetHyperlinks(new StepId("links"), "Budget", null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetComments(new StepId("comments"), "Budget", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadCommand.GetDataValidations(new StepId("validations"), "Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadCommand.GetAutofilters(new StepId("autofilters"), " "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetTables(new StepId("tables"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetPivotTables(new StepId("pivots"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetFormulaSurface(new StepId("formula"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookReadTestSupport.getSheetSchema("schema", "Budget", "A1", 0, 1));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.GetNamedRangeSurface(new StepId("surface"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.AnalyzeFormulaHealth(new StepId("formulaHealth"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadCommand.AnalyzeDataValidationHealth(
                new StepId("dataValidationHealth"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadCommand.AnalyzeConditionalFormattingHealth(
                new StepId("conditionalFormattingHealth"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadCommand.AnalyzeAutofilterHealth(new StepId("autofilterHealth"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.AnalyzeTableHealth(new StepId("tableHealth"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadCommand.AnalyzePivotTableHealth(new StepId("pivotTableHealth"), null));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadCommand.AnalyzeHyperlinkHealth(new StepId("hyperlinkHealth"), null));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookReadCommand.AnalyzeNamedRangeHealth(new StepId("namedRangeHealth"), null));
  }
}
