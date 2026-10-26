package dev.erst.gridgrind.core.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused coverage for selector factories and target-scoped fluent builders. */
class AuthoringTargetCoverageTest {
  @Test
  void targetFactoriesExposeTheExpectedSelectors() {
    assertInstanceOf(WorkbookSelector.Current.class, Targets.workbook().selector());
    assertInstanceOf(SheetSelector.ByName.class, Targets.sheet(new SheetName("Budget")).selector());
    assertInstanceOf(
        CellSelector.ByAddress.class,
        Targets.cell(new SheetName("Budget"), new CellAddress("A1")).selector());
    assertInstanceOf(
        RangeSelector.ByRange.class,
        Targets.range(new SheetName("Budget"), new RangeRef("A1:B2")).selector());
    assertInstanceOf(
        RangeSelector.RectangularWindow.class,
        Targets.window(new SheetName("Budget"), new CellAddress("A1"), 5, 3).selector());
    assertInstanceOf(
        TableSelector.ByName.class, Targets.table(new TableName("BudgetTable")).selector());
    assertInstanceOf(
        TableSelector.ByNameOnSheet.class,
        Targets.tableOnSheet(new TableName("BudgetTable"), new SheetName("Budget")).selector());
    assertInstanceOf(
        NamedRangeSelector.ByName.class,
        Targets.namedRange(new DefinedName("BudgetRange")).selector());
    assertInstanceOf(
        NamedRangeSelector.SheetScope.class,
        Targets.namedRangeOnSheet(new DefinedName("BudgetRange"), new SheetName("Budget"))
            .selector());
    assertInstanceOf(
        NamedRangeSelector.WorkbookScope.class,
        Targets.workbookNamedRange(new DefinedName("BudgetRange")).selector());
    assertInstanceOf(
        ChartSelector.ByName.class,
        Targets.chart(new SheetName("Budget"), "RevenueChart").selector());
    assertInstanceOf(
        PivotTableSelector.ByName.class, Targets.pivotTable("RevenuePivot").selector());
    assertInstanceOf(
        PivotTableSelector.ByNameOnSheet.class,
        Targets.pivotTableOnSheet("RevenuePivot", new SheetName("Budget")).selector());
  }

  @Test
  void workbookSheetCellRangeAndWindowTargetsCoverTheirSurface() {
    InspectionStep workbookSummary = Targets.workbook().summary().toStep(new StepId("summary"));
    InspectionStep workbookPackageSecurity =
        Targets.workbook().packageSecurity().toStep(new StepId("package-security"));
    InspectionStep workbookProtection =
        Targets.workbook().protection().toStep(new StepId("protection"));
    InspectionStep workbookFindings = Targets.workbook().findings().toStep(new StepId("findings"));
    assertInstanceOf(WorkbookIntrospectionQuery.GetWorkbookSummary.class, workbookSummary.query());
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetPackageSecurity.class, workbookPackageSecurity.query());
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetWorkbookProtection.class, workbookProtection.query());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeWorkbookFindings.class, workbookFindings.query());

    SheetTarget sheet = Targets.sheet(new SheetName("Budget"));
    assertInstanceOf(
        WorkbookMutationAction.EnsureSheet.class,
        sheet.ensureExists().toStep(new StepId("ensure")).action());
    assertInstanceOf(
        WorkbookMutationAction.RenameSheet.class,
        sheet.renameTo(new SheetName("Budget 2026")).toStep(new StepId("rename")).action());
    assertInstanceOf(
        WorkbookMutationAction.DeleteSheet.class,
        sheet.delete().toStep(new StepId("delete")).action());
    assertInstanceOf(
        WorkbookMutationAction.SetSheetZoom.class,
        sheet.setZoom(125).toStep(new StepId("zoom")).action());
    assertInstanceOf(
        WorkbookMutationAction.ClearPrintLayout.class,
        sheet.clearPrintLayout().toStep(new StepId("clear-print")).action());
    assertInstanceOf(
        SheetIntrospectionQuery.GetSheetSummary.class,
        sheet.summary().toStep(new StepId("summary")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetSheetLayout.class,
        sheet.layout().toStep(new StepId("layout")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetPrintLayout.class,
        sheet.printLayout().toStep(new StepId("print")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetMergedRegions.class,
        sheet.mergedRegions().toStep(new StepId("merged")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetAutofilters.class,
        sheet.autofilters().toStep(new StepId("autofilters")).query());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetCharts.class,
        sheet.charts().toStep(new StepId("charts")).query());
    InspectionStep drawingObjects = sheet.drawingObjects().toStep(new StepId("drawing-objects"));
    assertInstanceOf(DrawingObjectSelector.AllOnSheet.class, drawingObjects.target());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetDrawingObjects.class, drawingObjects.query());
    assertInstanceOf(
        InspectionSurfaceQuery.GetFormulaSurface.class,
        sheet.formulaSurface().toStep(new StepId("formula-surface")).query());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeFormulaHealth.class,
        sheet.formulaHealth().toStep(new StepId("formula-health")).query());

    CellTarget cell = Targets.cell(new SheetName("Budget"), new CellAddress("A1"));
    assertInstanceOf(
        CellMutationAction.SetCell.class,
        cell.set(Values.text("Owner")).toStep(new StepId("set")).action());
    assertInstanceOf(
        CellMutationAction.SetHyperlink.class,
        cell.setHyperlink(Links.url("https://example.com"))
            .toStep(new StepId("hyperlink"))
            .action());
    assertInstanceOf(
        CellMutationAction.ClearHyperlink.class,
        cell.clearHyperlink().toStep(new StepId("clear-link")).action());
    assertInstanceOf(
        CellMutationAction.SetComment.class,
        cell.setComment(Values.comment("note", "Ada")).toStep(new StepId("comment")).action());
    assertInstanceOf(
        CellMutationAction.ClearComment.class,
        cell.clearComment().toStep(new StepId("clear-comment")).action());
    assertInstanceOf(
        SheetIntrospectionQuery.GetCells.class, cell.read().toStep(new StepId("read")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetHyperlinks.class,
        cell.hyperlinks().toStep(new StepId("links")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetComments.class,
        cell.comments().toStep(new StepId("comments")).query());
    assertInstanceOf(
        CellAssertion.Value.class,
        cell.valueEquals(ExpectedValues.text("Owner")).toStep(new StepId("assert")).assertion());
    assertInstanceOf(
        CellAssertion.DisplayValue.class,
        cell.displayValueEquals("Owner").toStep(new StepId("display")).assertion());
    assertInstanceOf(
        CellAssertion.FormulaText.class,
        cell.formulaEquals("SUM(A1:A2)").toStep(new StepId("formula")).assertion());

    RangeTarget range = Targets.range(new SheetName("Budget"), new RangeRef("A1:B2"));
    MutationStep setRows =
        range
            .setRows(List.of(Values.row(Values.text("Owner"), Values.number(42.5))))
            .toStep(new StepId("rows"));
    assertInstanceOf(CellMutationAction.SetRange.class, setRows.action());
    assertInstanceOf(
        CellMutationAction.ClearRange.class, range.clear().toStep(new StepId("clear")).action());
    assertInstanceOf(
        WorkbookMutationAction.MergeCells.class,
        range.merge().toStep(new StepId("merge")).action());
    assertInstanceOf(
        WorkbookMutationAction.UnmergeCells.class,
        range.unmerge().toStep(new StepId("unmerge")).action());
    assertInstanceOf(
        SheetIntrospectionQuery.GetDataValidations.class,
        range.dataValidations().toStep(new StepId("validations")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetConditionalFormatting.class,
        range.conditionalFormatting().toStep(new StepId("formatting")).query());
    assertEquals(
        "rows must not be null",
        assertThrows(NullPointerException.class, () -> range.setRows(null)).getMessage());

    WindowTarget window = Targets.window(new SheetName("Budget"), new CellAddress("A1"), 5, 3);
    assertInstanceOf(
        SheetIntrospectionQuery.GetWindow.class, window.read().toStep(new StepId("read")).query());
    assertInstanceOf(
        InspectionSurfaceQuery.GetSheetSchema.class,
        window.schema().toStep(new StepId("schema")).query());
  }

  @Test
  void tableNamedRangeChartAndPivotTargetsCoverTheirSurface() {
    TableTarget table = Targets.table(new TableName("BudgetTable"));
    TableTarget tableOnSheet =
        Targets.tableOnSheet(new TableName("BudgetTable"), new SheetName("Budget"));
    assertInstanceOf(TableSelector.ByName.class, table.selector());
    assertInstanceOf(TableSelector.ByNameOnSheet.class, tableOnSheet.selector());
    assertInstanceOf(TableRowSelector.ByIndex.class, table.row(2).selector());
    assertInstanceOf(
        TableRowSelector.ByKeyCell.class,
        table.rowByKey("Item", Values.textFile(Path.of("authored-inputs", "item.txt"))).selector());
    assertInstanceOf(
        StructuredMutationAction.SetTable.class,
        tableOnSheet
            .define(
                Tables.define(
                    new TableName("BudgetTable"),
                    new SheetName("Budget"),
                    new RangeRef("A1:B3"),
                    false,
                    Tables.noStyle()))
            .toStep(new StepId("define"))
            .action());
    assertInstanceOf(
        StructuredMutationAction.DeleteTable.class,
        tableOnSheet.delete().toStep(new StepId("delete")).action());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetTables.class,
        table.inspect().toStep(new StepId("inspect")).query());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeTableHealth.class,
        table.analyzeHealth().toStep(new StepId("health")).query());
    assertInstanceOf(
        PresenceAssertion.TablePresent.class,
        table.present().toStep(new StepId("present")).assertion());
    assertInstanceOf(
        PresenceAssertion.TableAbsent.class,
        table.absent().toStep(new StepId("absent")).assertion());

    TableCellTarget tableCell = table.row(1).cell("Amount");
    assertInstanceOf(TableCellSelector.ByColumnName.class, tableCell.selector());
    assertInstanceOf(
        CellMutationAction.SetCell.class,
        tableCell.set(Values.number(125.0)).toStep(new StepId("set")).action());
    assertInstanceOf(
        CellMutationAction.SetHyperlink.class,
        tableCell.setHyperlink(Links.document("Budget!A1")).toStep(new StepId("link")).action());
    assertInstanceOf(
        CellMutationAction.ClearHyperlink.class,
        tableCell.clearHyperlink().toStep(new StepId("clear-link")).action());
    assertInstanceOf(
        CellMutationAction.SetComment.class,
        tableCell
            .setComment(Values.comment("raise", "Ada"))
            .toStep(new StepId("comment"))
            .action());
    assertInstanceOf(
        CellMutationAction.ClearComment.class,
        tableCell.clearComment().toStep(new StepId("clear-comment")).action());
    assertInstanceOf(
        SheetIntrospectionQuery.GetCells.class,
        tableCell.read().toStep(new StepId("read")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetHyperlinks.class,
        tableCell.hyperlinks().toStep(new StepId("links")).query());
    assertInstanceOf(
        SheetIntrospectionQuery.GetComments.class,
        tableCell.comments().toStep(new StepId("comments")).query());
    assertInstanceOf(
        CellAssertion.Value.class,
        tableCell
            .valueEquals(ExpectedValues.number(125.0))
            .toStep(new StepId("assert"))
            .assertion());
    assertInstanceOf(
        CellAssertion.DisplayValue.class,
        tableCell.displayValueEquals("125").toStep(new StepId("display")).assertion());
    assertInstanceOf(
        CellAssertion.FormulaText.class,
        tableCell.formulaEquals("SUM(A1:A2)").toStep(new StepId("formula")).assertion());

    NamedRangeTarget namedRange = Targets.namedRange(new DefinedName("BudgetRange"));
    NamedRangeTarget namedRangeOnSheet =
        Targets.namedRangeOnSheet(new DefinedName("BudgetRange"), new SheetName("Budget"));
    NamedRangeTarget workbookNamedRange =
        Targets.workbookNamedRange(new DefinedName("BudgetRange"));
    assertInstanceOf(NamedRangeSelector.ByName.class, namedRange.selector());
    assertInstanceOf(NamedRangeSelector.SheetScope.class, namedRangeOnSheet.selector());
    assertInstanceOf(NamedRangeSelector.WorkbookScope.class, workbookNamedRange.selector());
    assertInstanceOf(
        StructuredMutationAction.DeleteNamedRange.class,
        namedRange.delete().toStep(new StepId("delete")).action());
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetNamedRanges.class,
        namedRange.inspect().toStep(new StepId("inspect")).query());
    assertInstanceOf(
        InspectionSurfaceQuery.GetNamedRangeSurface.class,
        namedRange.surface().toStep(new StepId("surface")).query());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeNamedRangeHealth.class,
        namedRange.analyzeHealth().toStep(new StepId("health")).query());
    assertInstanceOf(
        PresenceAssertion.NamedRangePresent.class,
        namedRange.present().toStep(new StepId("present")).assertion());
    assertInstanceOf(
        PresenceAssertion.NamedRangeAbsent.class,
        namedRange.absent().toStep(new StepId("absent")).assertion());

    ChartTarget chart = Targets.chart(new SheetName("Budget"), "RevenueChart");
    ChartTarget allChartsOnSheet =
        new ChartTarget(new ChartSelector.AllOnSheet(new SheetName("Budget")));
    InspectionStep namedChartInspection = chart.inspect().toStep(new StepId("named-chart"));
    InspectionStep allChartsInspection =
        allChartsOnSheet.inspect().toStep(new StepId("all-charts"));
    assertInstanceOf(ChartSelector.ByName.class, namedChartInspection.target());
    assertInstanceOf(ChartSelector.AllOnSheet.class, allChartsInspection.target());
    assertInstanceOf(WorkbookAssetIntrospectionQuery.GetCharts.class, namedChartInspection.query());
    assertInstanceOf(WorkbookAssetIntrospectionQuery.GetCharts.class, allChartsInspection.query());
    assertInstanceOf(
        PresenceAssertion.ChartPresent.class,
        chart.present().toStep(new StepId("present")).assertion());
    assertInstanceOf(
        PresenceAssertion.ChartAbsent.class,
        chart.absent().toStep(new StepId("absent")).assertion());

    PivotTableTarget pivotTable = Targets.pivotTable("RevenuePivot");
    PivotTableTarget pivotTableOnSheet =
        Targets.pivotTableOnSheet("RevenuePivot", new SheetName("Budget"));
    assertInstanceOf(PivotTableSelector.ByName.class, pivotTable.selector());
    assertInstanceOf(PivotTableSelector.ByNameOnSheet.class, pivotTableOnSheet.selector());
    assertInstanceOf(
        StructuredMutationAction.DeletePivotTable.class,
        pivotTableOnSheet.delete().toStep(new StepId("delete")).action());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetPivotTables.class,
        pivotTable.inspect().toStep(new StepId("inspect")).query());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzePivotTableHealth.class,
        pivotTable.analyzeHealth().toStep(new StepId("health")).query());
    assertInstanceOf(
        PresenceAssertion.PivotTablePresent.class,
        pivotTable.present().toStep(new StepId("present")).assertion());
    assertInstanceOf(
        PresenceAssertion.PivotTableAbsent.class,
        pivotTable.absent().toStep(new StepId("absent")).assertion());
  }

  @Test
  void targetRefsRejectNullSelectorsImmediately() {
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new WorkbookTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new SheetTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new CellTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new RangeTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new WindowTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new TableTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new TableRowTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new TableCellTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new NamedRangeTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new ChartTarget(null)).getMessage());
    assertEquals(
        "selector must not be null",
        assertThrows(NullPointerException.class, () -> new PivotTableTarget(null)).getMessage());
  }
}
