package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for step validation seams introduced by the canonical step envelope. */
class WorkbookStepValidationTest {
  @Test
  void validatesStepIdsAndTargets() {
    assertEquals("step-01", WorkbookStepValidation.requireStepId("step-01"));
    assertEquals(
        new WorkbookSelector.Current(),
        WorkbookStepValidation.requireTarget(new WorkbookSelector.Current()));

    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> WorkbookStepValidation.requireStepId(" "))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(NullPointerException.class, () -> WorkbookStepValidation.requireTarget(null))
            .getMessage());
    assertEquals(
        "action must not be null",
        assertThrows(
                NullPointerException.class,
                () -> WorkbookStepValidation.allowedTargetTypes((MutationAction) null))
            .getMessage());
    assertEquals(
        "query must not be null",
        assertThrows(
                NullPointerException.class,
                () -> WorkbookStepValidation.allowedTargetTypes((InspectionQuery) null))
            .getMessage());
  }

  @Test
  void validatesCompatibleMutationTargetsAcrossSingleAndUnionTargetFamilies() {
    Selector setCellTarget = new CellSelector.ByAddress("Budget", "A1");
    MutationAction setCell = new MutationAction.SetCell(new CellInput.Text(text("Owner")));
    assertEquals(setCell, WorkbookStepValidation.requireCompatible(setCellTarget, setCell));

    Selector tableTarget = new TableSelector.ByName("BudgetTable");
    MutationAction deleteTable = new MutationAction.DeleteTable();
    assertEquals(deleteTable, WorkbookStepValidation.requireCompatible(tableTarget, deleteTable));

    IllegalArgumentException wrongTarget =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.requireCompatible(
                    new SheetSelector.ByName("Budget"), setCell));
    assertEquals(
        "SET_CELL requires target type ByAddress or ByColumnName but got ByName",
        wrongTarget.getMessage());

    IllegalArgumentException unionTargetFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.requireCompatible(
                    new WorkbookSelector.Current(), deleteTable));
    assertEquals(
        "DELETE_TABLE requires target type ByNameOnSheet or ByName but got Current",
        unionTargetFailure.getMessage());
  }

  @Test
  void validatesCompatibleInspectionTargetsAcrossSingleAndUnionTargetFamilies() {
    InspectionQuery getCharts = new InspectionQuery.GetCharts();
    assertEquals(
        getCharts,
        WorkbookStepValidation.requireCompatible(
            new ChartSelector.AllOnSheet("Budget"), getCharts));

    InspectionQuery analyzeNamedRangeHealth = new InspectionQuery.AnalyzeNamedRangeHealth();
    assertEquals(
        analyzeNamedRangeHealth,
        WorkbookStepValidation.requireCompatible(
            new NamedRangeSelector.WorkbookScope("BudgetTotal"), analyzeNamedRangeHealth));

    IllegalArgumentException wrongTarget =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.requireCompatible(
                    new RangeSelector.ByRange("Budget", "A1:B2"), getCharts));
    assertEquals(
        "GET_CHARTS requires target type AllOnSheet or ByName but got ByRange",
        wrongTarget.getMessage());
  }

  @Test
  void validatesCompatibleAssertionTargetsAcrossDirectAnalysisAndCompositeFamilies() {
    Assertion cellValue = new Assertion.CellValue(new ExpectedCellValue.Text("Owner"));
    assertEquals(
        cellValue,
        WorkbookStepValidation.requireCompatible(
            new CellSelector.ByAddress("Budget", "A1"), cellValue));

    Assertion analysisSeverity =
        new Assertion.AnalysisMaxSeverity(
            new InspectionQuery.AnalyzeFormulaHealth(),
            dev.erst.gridgrind.contract.dto.AnalysisSeverity.WARNING);
    assertEquals(
        analysisSeverity,
        WorkbookStepValidation.requireCompatible(
            new SheetSelector.ByName("Budget"), analysisSeverity));

    Assertion anyOf = new Assertion.AnyOf(List.of(new Assertion.Present(), new Assertion.Absent()));
    assertEquals(
        anyOf,
        WorkbookStepValidation.requireCompatible(
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"), anyOf));

    IllegalArgumentException wrongTarget =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.requireCompatible(
                    new WorkbookSelector.Current(),
                    new Assertion.CellValue(new ExpectedCellValue.Text("Owner"))));
    assertEquals(
        "EXPECT_CELL_VALUE requires target type ByAddress, ByAddresses or ByColumnName but got Current",
        wrongTarget.getMessage());

    IllegalArgumentException incompatibleComposite =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.allowedTargetTypes(
                    new Assertion.AllOf(
                        List.of(
                            new Assertion.CellValue(new ExpectedCellValue.Text("Owner")),
                            new Assertion.Present()))));
    assertEquals(
        "ALL_OF requires nested assertions with compatible target families",
        incompatibleComposite.getMessage());
  }

  @Test
  void exposesAllowedSelectorTypeFamiliesForStepQueriesAndActions() {
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.EnsureSheet())));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class, TableSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeleteTable())));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetWorkbookSummary())));
    assertEquals(
        List.of(ChartSelector.AllOnSheet.class, SheetSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetCharts())));
    assertEquals(
        List.of(
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new Assertion.CellValue(new ExpectedCellValue.Text("Owner")))));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new Assertion.AnalysisFindingPresent(
                    new InspectionQuery.AnalyzeFormulaHealth(),
                    dev.erst.gridgrind.contract.dto.AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    null,
                    null))));
  }

  @Test
  void exposesEveryRemainingSelectorFamilyBranchAcrossActionsAndQueries() {
    assertEquals(
        List.of(RowBandSelector.Insertion.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.InsertRows())));
    assertEquals(
        List.of(RowBandSelector.Span.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.GroupRows(null))));
    assertEquals(
        List.of(ColumnBandSelector.Insertion.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.InsertColumns())));
    assertEquals(
        List.of(ColumnBandSelector.Span.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.GroupColumns(null))));
    assertEquals(
        List.of(DrawingObjectSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeleteDrawingObject())));
    assertEquals(
        List.of(PivotTableSelector.ByNameOnSheet.class, PivotTableSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeletePivotTable())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.AnalyzeConditionalFormattingHealth())));
    assertEquals(
        List.of(
            CellSelector.AllUsedInSheet.class,
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetComments())));
    assertEquals(
        List.of(RangeSelector.RectangularWindow.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetSheetSchema())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.GetConditionalFormatting())));
    assertEquals(
        List.of(DrawingObjectSelector.AllOnSheet.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetDrawingObjects())));
    assertEquals(
        List.of(DrawingObjectSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.GetDrawingObjectPayload())));
    assertEquals(
        List.of(PivotTableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.AnalyzePivotTableHealth())));
    assertEquals(
        List.of(TableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.AnalyzeTableHealth())));
  }

  @Test
  void coversEveryRemainingGroupedActionAndQueryCaseAndThreeWayUnionWording() {
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.SetPicture(
                    new dev.erst.gridgrind.contract.dto.PictureInput(
                        "Logo",
                        new dev.erst.gridgrind.contract.dto.PictureDataInput(
                            dev.erst.gridgrind.excel.ExcelPictureFormat.PNG, binary("AQID")),
                        new dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell(
                            new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(0, 0, 0, 0),
                            new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(1, 1, 0, 0),
                            null),
                        null)))));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.SetWorkbookProtection(
                    new dev.erst.gridgrind.contract.dto.WorkbookProtectionInput(
                        true, false, false, null, null)))));
    assertEquals(
        List.of(RangeSelector.ByRange.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.MergeCells())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new MutationAction.ClearDataValidations())));
    assertEquals(
        List.of(ColumnBandSelector.Span.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new MutationAction.SetColumnWidth(8.43d))));
    assertEquals(
        List.of(RowBandSelector.Span.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.SetRowHeight(15.0d))));
    assertEquals(
        List.of(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.ClearHyperlink())));
    assertEquals(
        List.of(DrawingObjectSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.SetDrawingObjectAnchor(
                    new dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell(
                        new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(0, 0, 0, 0),
                        new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(1, 1, 0, 0),
                        null)))));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class, TableSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.SetTable(
                    new dev.erst.gridgrind.contract.dto.TableInput(
                        "BudgetTable",
                        "Budget",
                        "A1:B2",
                        false,
                        new dev.erst.gridgrind.contract.dto.TableStyleInput.None())))));
    assertEquals(
        List.of(PivotTableSelector.ByNameOnSheet.class, PivotTableSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.SetPivotTable(
                    new dev.erst.gridgrind.contract.dto.PivotTableInput(
                        "Pivot",
                        "Report",
                        new dev.erst.gridgrind.contract.dto.PivotTableInput.Source.Range(
                            "Budget", "A1:B2"),
                        new dev.erst.gridgrind.contract.dto.PivotTableInput.Anchor("A3"),
                        List.of("Category"),
                        List.of(),
                        List.of(),
                        List.of(
                            new dev.erst.gridgrind.contract.dto.PivotTableInput.DataField(
                                "Amount",
                                dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction.SUM,
                                "Total Amount",
                                null)))))));
    assertEquals(
        List.of(
            NamedRangeSelector.ByName.class,
            NamedRangeSelector.WorkbookScope.class,
            NamedRangeSelector.SheetScope.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.SetNamedRange(
                    "BudgetTotal",
                    new dev.erst.gridgrind.contract.dto.NamedRangeScope.Workbook(),
                    new dev.erst.gridgrind.contract.dto.NamedRangeTarget("Budget", "A1")))));

    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.AnalyzeWorkbookFindings())));
    assertEquals(
        List.of(NamedRangeSelector.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetNamedRanges())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetSheetLayout())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetFormulaSurface())));
    assertEquals(
        List.of(
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetCells())));
    assertEquals(
        List.of(RangeSelector.RectangularWindow.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetWindow())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetDataValidations())));
    assertEquals(
        List.of(PivotTableSelector.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetPivotTables())));
    assertEquals(
        List.of(TableSelector.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetTables())));

    IllegalArgumentException threeWayUnionFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.requireCompatible(
                    new WorkbookSelector.Current(), new MutationAction.DeleteNamedRange()));
    assertEquals(
        "DELETE_NAMED_RANGE requires target type ByName, WorkbookScope or SheetScope but got Current",
        threeWayUnionFailure.getMessage());

    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.AutoSizeColumns())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.ClearAutofilter())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.AppendRow(List.of(new CellInput.Text(text("A")))))));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.ClearWorkbookProtection())));
    assertEquals(
        List.of(RangeSelector.ByRange.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.SetRange(List.of(List.of(new CellInput.Text(text("A"))))))));
    assertEquals(
        List.of(RangeSelector.ByRange.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.UnmergeCells())));
    assertEquals(
        List.of(RangeSelector.ByRange.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.ClearRange())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.ClearConditionalFormatting())));
    assertEquals(
        List.of(ColumnBandSelector.Span.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeleteColumns())));
    assertEquals(
        List.of(RowBandSelector.Span.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeleteRows())));
    assertEquals(
        List.of(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.ClearComment())));
    assertEquals(
        List.of(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new MutationAction.SetHyperlink(
                    new dev.erst.gridgrind.contract.dto.HyperlinkTarget.Url(
                        "https://example.com")))));
    assertEquals(
        List.of(DrawingObjectSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeleteDrawingObject())));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class, TableSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeleteTable())));
    assertEquals(
        List.of(PivotTableSelector.ByNameOnSheet.class, PivotTableSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeletePivotTable())));
    assertEquals(
        List.of(
            NamedRangeSelector.ByName.class,
            NamedRangeSelector.WorkbookScope.class,
            NamedRangeSelector.SheetScope.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new MutationAction.DeleteNamedRange())));

    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.GetWorkbookProtection())));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetPackageSecurity())));
    assertEquals(
        List.of(NamedRangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetNamedRangeSurface())));
    assertEquals(
        List.of(NamedRangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.AnalyzeNamedRangeHealth())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetAutofilters())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetMergedRegions())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetPrintLayout())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.AnalyzeHyperlinkHealth())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.AnalyzeFormulaHealth())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.AnalyzeDataValidationHealth())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.AnalyzeAutofilterHealth())));
    assertEquals(
        List.of(
            CellSelector.AllUsedInSheet.class,
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetHyperlinks())));
    assertEquals(
        List.of(
            CellSelector.AllUsedInSheet.class,
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetComments())));
    assertEquals(
        List.of(RangeSelector.RectangularWindow.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.GetSheetSchema())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.GetConditionalFormatting())));
    assertEquals(
        List.of(PivotTableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionQuery.AnalyzePivotTableHealth())));
    assertEquals(
        List.of(TableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new InspectionQuery.AnalyzeTableHealth())));
  }

  @Test
  void acceptsTableCellSelectorsOnlyForExactCellStepFamilies() {
    TableCellSelector.ByColumnName tableCell =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(text("Hosting"))),
            "Amount");

    assertEquals(
        new MutationAction.SetCell(new CellInput.Numeric(125.0)),
        WorkbookStepValidation.requireCompatible(
            tableCell, new MutationAction.SetCell(new CellInput.Numeric(125.0))));
    assertEquals(
        new InspectionQuery.GetCells(),
        WorkbookStepValidation.requireCompatible(tableCell, new InspectionQuery.GetCells()));
    assertEquals(
        new Assertion.DisplayValue("125"),
        WorkbookStepValidation.requireCompatible(tableCell, new Assertion.DisplayValue("125")));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }

  private static BinarySourceInput binary(String value) {
    return BinarySourceInput.inlineBase64(value);
  }
}
