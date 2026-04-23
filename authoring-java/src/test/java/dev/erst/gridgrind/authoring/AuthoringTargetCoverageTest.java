package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused coverage for selector factories and target-scoped fluent builders. */
class AuthoringTargetCoverageTest {
  @Test
  void targetFactoriesExposeTheExpectedSelectors() {
    assertInstanceOf(WorkbookSelector.Current.class, Targets.workbook().selector());
    assertInstanceOf(SheetSelector.ByName.class, Targets.sheet("Budget").selector());
    assertInstanceOf(CellSelector.ByAddress.class, Targets.cell("Budget", "A1").selector());
    assertInstanceOf(RangeSelector.ByRange.class, Targets.range("Budget", "A1:B2").selector());
    assertInstanceOf(
        RangeSelector.RectangularWindow.class, Targets.window("Budget", "A1", 5, 3).selector());
    assertInstanceOf(TableSelector.ByName.class, Targets.table("BudgetTable").selector());
    assertInstanceOf(
        TableSelector.ByNameOnSheet.class,
        Targets.tableOnSheet("BudgetTable", "Budget").selector());
    assertInstanceOf(NamedRangeSelector.ByName.class, Targets.namedRange("BudgetRange").selector());
    assertInstanceOf(
        NamedRangeSelector.SheetScope.class,
        Targets.namedRangeOnSheet("BudgetRange", "Budget").selector());
    assertInstanceOf(
        NamedRangeSelector.WorkbookScope.class,
        Targets.workbookNamedRange("BudgetRange").selector());
    assertInstanceOf(
        ChartSelector.ByName.class, Targets.chart("Budget", "RevenueChart").selector());
    assertInstanceOf(
        PivotTableSelector.ByName.class, Targets.pivotTable("RevenuePivot").selector());
    assertInstanceOf(
        PivotTableSelector.ByNameOnSheet.class,
        Targets.pivotTableOnSheet("RevenuePivot", "Budget").selector());
  }

  @Test
  void workbookSheetCellRangeAndWindowTargetsCoverTheirSurface() {
    InspectionStep workbookSummary = Targets.workbook().summary().toStep("summary");
    InspectionStep workbookPackageSecurity =
        Targets.workbook().packageSecurity().toStep("package-security");
    InspectionStep workbookProtection = Targets.workbook().protection().toStep("protection");
    InspectionStep workbookFindings = Targets.workbook().findings().toStep("findings");
    assertInstanceOf(InspectionQuery.GetWorkbookSummary.class, workbookSummary.query());
    assertInstanceOf(InspectionQuery.GetPackageSecurity.class, workbookPackageSecurity.query());
    assertInstanceOf(InspectionQuery.GetWorkbookProtection.class, workbookProtection.query());
    assertInstanceOf(InspectionQuery.AnalyzeWorkbookFindings.class, workbookFindings.query());

    Targets.SheetRef sheet = Targets.sheet("Budget");
    assertInstanceOf(
        MutationAction.EnsureSheet.class, sheet.ensureExists().toStep("ensure").action());
    assertInstanceOf(
        MutationAction.RenameSheet.class, sheet.renameTo("Budget 2026").toStep("rename").action());
    assertInstanceOf(MutationAction.DeleteSheet.class, sheet.delete().toStep("delete").action());
    assertInstanceOf(MutationAction.SetSheetZoom.class, sheet.setZoom(125).toStep("zoom").action());
    assertInstanceOf(
        MutationAction.ClearPrintLayout.class,
        sheet.clearPrintLayout().toStep("clear-print").action());
    assertInstanceOf(
        InspectionQuery.GetSheetSummary.class, sheet.summary().toStep("summary").query());
    assertInstanceOf(InspectionQuery.GetSheetLayout.class, sheet.layout().toStep("layout").query());
    assertInstanceOf(
        InspectionQuery.GetPrintLayout.class, sheet.printLayout().toStep("print").query());
    assertInstanceOf(
        InspectionQuery.GetMergedRegions.class, sheet.mergedRegions().toStep("merged").query());
    assertInstanceOf(
        InspectionQuery.GetAutofilters.class, sheet.autofilters().toStep("autofilters").query());
    assertInstanceOf(InspectionQuery.GetCharts.class, sheet.charts().toStep("charts").query());
    InspectionStep drawingObjects = sheet.drawingObjects().toStep("drawing-objects");
    assertInstanceOf(DrawingObjectSelector.AllOnSheet.class, drawingObjects.target());
    assertInstanceOf(InspectionQuery.GetDrawingObjects.class, drawingObjects.query());
    assertInstanceOf(
        InspectionQuery.GetFormulaSurface.class,
        sheet.formulaSurface().toStep("formula-surface").query());
    assertInstanceOf(
        InspectionQuery.AnalyzeFormulaHealth.class,
        sheet.formulaHealth().toStep("formula-health").query());

    Targets.CellRef cell = Targets.cell("Budget", "A1");
    assertInstanceOf(
        MutationAction.SetCell.class, cell.set(Values.text("Owner")).toStep("set").action());
    assertInstanceOf(
        MutationAction.SetHyperlink.class,
        cell.setHyperlink(Links.url("https://example.com")).toStep("hyperlink").action());
    assertInstanceOf(
        MutationAction.ClearHyperlink.class, cell.clearHyperlink().toStep("clear-link").action());
    assertInstanceOf(
        MutationAction.SetComment.class,
        cell.setComment(Values.comment("note", "Ada")).toStep("comment").action());
    assertInstanceOf(
        MutationAction.ClearComment.class, cell.clearComment().toStep("clear-comment").action());
    assertInstanceOf(InspectionQuery.GetCells.class, cell.read().toStep("read").query());
    assertInstanceOf(
        InspectionQuery.GetHyperlinks.class, cell.hyperlinks().toStep("links").query());
    assertInstanceOf(InspectionQuery.GetComments.class, cell.comments().toStep("comments").query());
    assertInstanceOf(
        Assertion.CellValue.class,
        cell.valueEquals(Values.expectedText("Owner")).toStep("assert").assertion());
    assertInstanceOf(
        Assertion.DisplayValue.class,
        cell.displayValueEquals("Owner").toStep("display").assertion());
    assertInstanceOf(
        Assertion.FormulaText.class,
        cell.formulaEquals("SUM(A1:A2)").toStep("formula").assertion());

    Targets.RangeRef range = Targets.range("Budget", "A1:B2");
    MutationStep setRows =
        range
            .setRows(List.of(Values.row(Values.text("Owner"), Values.number(42.5))))
            .toStep("rows");
    assertInstanceOf(MutationAction.SetRange.class, setRows.action());
    assertInstanceOf(MutationAction.ClearRange.class, range.clear().toStep("clear").action());
    assertInstanceOf(MutationAction.MergeCells.class, range.merge().toStep("merge").action());
    assertInstanceOf(MutationAction.UnmergeCells.class, range.unmerge().toStep("unmerge").action());
    assertInstanceOf(
        InspectionQuery.GetDataValidations.class,
        range.dataValidations().toStep("validations").query());
    assertInstanceOf(
        InspectionQuery.GetConditionalFormatting.class,
        range.conditionalFormatting().toStep("formatting").query());
    assertEquals(
        "rows must not be null",
        assertThrows(NullPointerException.class, () -> range.setRows(null)).getMessage());

    Targets.WindowRef window = Targets.window("Budget", "A1", 5, 3);
    assertInstanceOf(InspectionQuery.GetWindow.class, window.read().toStep("read").query());
    assertInstanceOf(
        InspectionQuery.GetSheetSchema.class, window.schema().toStep("schema").query());
  }

  @Test
  void tableNamedRangeChartAndPivotTargetsCoverTheirSurface() {
    Targets.TableRef table = Targets.table("BudgetTable");
    Targets.TableRef tableOnSheet = Targets.tableOnSheet("BudgetTable", "Budget");
    assertInstanceOf(TableSelector.ByName.class, table.selector());
    assertInstanceOf(TableSelector.ByNameOnSheet.class, tableOnSheet.selector());
    assertInstanceOf(TableRowSelector.ByIndex.class, table.row(2).selector());
    assertInstanceOf(
        TableRowSelector.ByKeyCell.class,
        table.rowByKey("Item", Values.textFile(Path.of("authored-inputs", "item.txt"))).selector());
    assertInstanceOf(
        MutationAction.SetTable.class,
        tableOnSheet
            .define(Tables.define("BudgetTable", "Budget", "A1:B3", false, Tables.noStyle()))
            .toStep("define")
            .action());
    assertInstanceOf(
        MutationAction.DeleteTable.class, tableOnSheet.delete().toStep("delete").action());
    assertInstanceOf(InspectionQuery.GetTables.class, table.inspect().toStep("inspect").query());
    assertInstanceOf(
        InspectionQuery.AnalyzeTableHealth.class, table.analyzeHealth().toStep("health").query());
    assertInstanceOf(Assertion.Present.class, table.present().toStep("present").assertion());
    assertInstanceOf(Assertion.Absent.class, table.absent().toStep("absent").assertion());

    Targets.TableCellRef tableCell = table.row(1).cell("Amount");
    assertInstanceOf(TableCellSelector.ByColumnName.class, tableCell.selector());
    assertInstanceOf(
        MutationAction.SetCell.class, tableCell.set(Values.number(125.0)).toStep("set").action());
    assertInstanceOf(
        MutationAction.SetHyperlink.class,
        tableCell.setHyperlink(Links.document("Budget!A1")).toStep("link").action());
    assertInstanceOf(
        MutationAction.ClearHyperlink.class,
        tableCell.clearHyperlink().toStep("clear-link").action());
    assertInstanceOf(
        MutationAction.SetComment.class,
        tableCell.setComment(Values.comment("raise", "Ada")).toStep("comment").action());
    assertInstanceOf(
        MutationAction.ClearComment.class,
        tableCell.clearComment().toStep("clear-comment").action());
    assertInstanceOf(InspectionQuery.GetCells.class, tableCell.read().toStep("read").query());
    assertInstanceOf(
        InspectionQuery.GetHyperlinks.class, tableCell.hyperlinks().toStep("links").query());
    assertInstanceOf(
        InspectionQuery.GetComments.class, tableCell.comments().toStep("comments").query());
    assertInstanceOf(
        Assertion.CellValue.class,
        tableCell.valueEquals(Values.expectedNumber(125.0)).toStep("assert").assertion());
    assertInstanceOf(
        Assertion.DisplayValue.class,
        tableCell.displayValueEquals("125").toStep("display").assertion());
    assertInstanceOf(
        Assertion.FormulaText.class,
        tableCell.formulaEquals("SUM(A1:A2)").toStep("formula").assertion());

    Targets.NamedRangeRef namedRange = Targets.namedRange("BudgetRange");
    Targets.NamedRangeRef namedRangeOnSheet = Targets.namedRangeOnSheet("BudgetRange", "Budget");
    Targets.NamedRangeRef workbookNamedRange = Targets.workbookNamedRange("BudgetRange");
    assertInstanceOf(NamedRangeSelector.ByName.class, namedRange.selector());
    assertInstanceOf(NamedRangeSelector.SheetScope.class, namedRangeOnSheet.selector());
    assertInstanceOf(NamedRangeSelector.WorkbookScope.class, workbookNamedRange.selector());
    assertInstanceOf(
        MutationAction.DeleteNamedRange.class, namedRange.delete().toStep("delete").action());
    assertInstanceOf(
        InspectionQuery.GetNamedRanges.class, namedRange.inspect().toStep("inspect").query());
    assertInstanceOf(
        InspectionQuery.GetNamedRangeSurface.class, namedRange.surface().toStep("surface").query());
    assertInstanceOf(
        InspectionQuery.AnalyzeNamedRangeHealth.class,
        namedRange.analyzeHealth().toStep("health").query());
    assertInstanceOf(Assertion.Present.class, namedRange.present().toStep("present").assertion());
    assertInstanceOf(Assertion.Absent.class, namedRange.absent().toStep("absent").assertion());

    Targets.ChartRef chart = Targets.chart("Budget", "RevenueChart");
    Targets.ChartRef allChartsOnSheet =
        new Targets.ChartRef(new ChartSelector.AllOnSheet("Budget"));
    InspectionStep namedChartInspection = chart.inspectOnSheet().toStep("named-chart");
    InspectionStep allChartsInspection = allChartsOnSheet.inspectOnSheet().toStep("all-charts");
    assertInstanceOf(SheetSelector.ByName.class, namedChartInspection.target());
    assertInstanceOf(SheetSelector.ByName.class, allChartsInspection.target());
    assertInstanceOf(InspectionQuery.GetCharts.class, namedChartInspection.query());
    assertInstanceOf(InspectionQuery.GetCharts.class, allChartsInspection.query());
    assertInstanceOf(Assertion.Present.class, chart.present().toStep("present").assertion());
    assertInstanceOf(Assertion.Absent.class, chart.absent().toStep("absent").assertion());

    Targets.PivotTableRef pivotTable = Targets.pivotTable("RevenuePivot");
    Targets.PivotTableRef pivotTableOnSheet = Targets.pivotTableOnSheet("RevenuePivot", "Budget");
    assertInstanceOf(PivotTableSelector.ByName.class, pivotTable.selector());
    assertInstanceOf(PivotTableSelector.ByNameOnSheet.class, pivotTableOnSheet.selector());
    assertInstanceOf(
        MutationAction.DeletePivotTable.class,
        pivotTableOnSheet.delete().toStep("delete").action());
    assertInstanceOf(
        InspectionQuery.GetPivotTables.class, pivotTable.inspect().toStep("inspect").query());
    assertInstanceOf(
        InspectionQuery.AnalyzePivotTableHealth.class,
        pivotTable.analyzeHealth().toStep("health").query());
    assertInstanceOf(Assertion.Present.class, pivotTable.present().toStep("present").assertion());
    assertInstanceOf(Assertion.Absent.class, pivotTable.absent().toStep("absent").assertion());
  }

  @Test
  void targetRefsRejectNullSelectorsImmediately() {
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.WorkbookRef(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.SheetRef(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.CellRef(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.RangeRef(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.WindowRef(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.TableRef(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.TableRowRef(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.TableCellRef(null))
            .getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.NamedRangeRef(null))
            .getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.ChartRef(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new Targets.PivotTableRef(null))
            .getMessage());
  }
}
