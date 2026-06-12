package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.query.*;
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
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Optional;
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
        "stepId must match [A-Za-z0-9._-]+",
        assertThrows(
                IllegalArgumentException.class,
                () -> WorkbookStepValidation.requireStepId("step with spaces"))
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
    MutationAction setCell = new CellMutationAction.SetCell(new CellInput.Text(text("Owner")));
    assertEquals(setCell, WorkbookStepValidation.requireCompatible(setCellTarget, setCell));

    Selector tableTarget = new TableSelector.ByNameOnSheet("BudgetTable", "Budget");
    MutationAction deleteTable = new StructuredMutationAction.DeleteTable();
    assertEquals(deleteTable, WorkbookStepValidation.requireCompatible(tableTarget, deleteTable));

    IllegalArgumentException wrongTarget =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.requireCompatible(
                    new SheetSelector.ByName("Budget"), setCell));
    assertEquals(
        "SET_CELL requires target type CELL_BY_ADDRESS or TABLE_CELL_BY_COLUMN_NAME but got"
            + " SHEET_BY_NAME",
        wrongTarget.getMessage());

    IllegalArgumentException unionTargetFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.requireCompatible(
                    new WorkbookSelector.Current(), deleteTable));
    assertEquals(
        "DELETE_TABLE requires target type TABLE_BY_NAME_ON_SHEET but got WORKBOOK_CURRENT",
        unionTargetFailure.getMessage());
  }

  @Test
  void validatesCompatibleInspectionTargetsAcrossSingleAndUnionTargetFamilies() {
    InspectionQuery getCharts = new WorkbookAssetIntrospectionQuery.GetCharts();
    assertEquals(
        getCharts,
        WorkbookStepValidation.requireCompatible(
            new ChartSelector.AllOnSheet("Budget"), getCharts));

    InspectionQuery analyzeNamedRangeHealth = new InspectionAnalysisQuery.AnalyzeNamedRangeHealth();
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
        "GET_CHARTS requires target type CHART_ALL_ON_SHEET or CHART_BY_NAME but got"
            + " RANGE_BY_RANGE",
        wrongTarget.getMessage());
  }

  @Test
  void validatesCompatibleAssertionTargetsAcrossDirectAnalysisAndCompositeFamilies() {
    Assertion cellValue =
        new CellAssertion.CellValue(
            new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner"));
    assertEquals(
        cellValue,
        WorkbookStepValidation.requireCompatible(
            new CellSelector.ByAddress("Budget", "A1"), cellValue));

    Assertion analysisSeverity =
        new AnalysisAssertion.AnalysisMaxSeverity(
            new InspectionAnalysisQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING);
    assertEquals(
        analysisSeverity,
        WorkbookStepValidation.requireCompatible(
            new SheetSelector.ByName("Budget"), analysisSeverity));

    Assertion anyOf =
        new CompositeAssertion.AnyOf(
            List.of(new PresenceAssertion.TablePresent(), new PresenceAssertion.TableAbsent()));
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
                    new CellAssertion.CellValue(
                        new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner"))));
    assertEquals(
        "EXPECT_CELL_VALUE requires target type CELL_BY_ADDRESS, CELL_BY_ADDRESSES or"
            + " TABLE_CELL_BY_COLUMN_NAME but got WORKBOOK_CURRENT",
        wrongTarget.getMessage());

    IllegalArgumentException incompatibleComposite =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.allowedTargetTypes(
                    new CompositeAssertion.AllOf(
                        List.of(
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner")),
                            new PresenceAssertion.TablePresent()))));
    assertEquals(
        "ALL_OF requires nested assertions with compatible target families",
        incompatibleComposite.getMessage());
  }

  @Test
  void exposesAllowedSelectorTypeFamiliesForStepQueriesAndActions() {
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new WorkbookMutationAction.EnsureSheet())));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new StructuredMutationAction.DeleteTable())));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.ImportCustomXmlMapping(
                    new dev.erst.gridgrind.contract.dto.CustomXmlImportInput(
                        new dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator(1L, "Map"),
                        new TextSourceInput.Inline("<root/>"))))));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookIntrospectionQuery.GetCustomXmlMappings())));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                    new dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator(1L, "Map"),
                    false,
                    "UTF-8"))));
    assertEquals(
        List.of(ChartSelector.AllOnSheet.class, ChartSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookAssetIntrospectionQuery.GetCharts())));
    assertEquals(
        List.of(
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new CellAssertion.CellValue(
                    new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner")))));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new AnalysisAssertion.AnalysisFindingPresent(
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    Optional.empty(),
                    Optional.empty()))));
  }

  @Test
  void rejectsStaticSelectorLookupForDynamicAssertionFamilies() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.staticAllowedTargetTypesForAssertionType(
                    AnalysisAssertion.AnalysisFindingPresent.class));

    assertEquals(
        "Assertion type "
            + AnalysisAssertion.AnalysisFindingPresent.class.getName()
            + " derives target selectors dynamically: Matches the nested analysis query's target selectors.",
        failure.getMessage());
  }

  @Test
  void exposesAssertionTargetSelectorRulesForDynamicAndDirectSelectorFamilies() {
    assertEquals(
        Optional.of("Matches the nested analysis query's target selectors."),
        WorkbookStepValidation.targetSelectorRuleForAssertionType(
            AnalysisAssertion.AnalysisFindingAbsent.class));
    assertEquals(
        Optional.empty(),
        WorkbookStepValidation.targetSelectorRuleForAssertionType(CellAssertion.CellValue.class));
    assertEquals(
        Optional.empty(),
        WorkbookStepValidation.targetSelectorRuleForAssertionType(
            PresenceAssertion.TableAbsent.class));
  }

  @Test
  void exposesEveryRemainingSelectorFamilyBranchAcrossActionsAndQueries() {
    assertEquals(
        List.of(RowBandSelector.Insertion.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new WorkbookMutationAction.InsertRows())));
    assertEquals(
        List.of(RowBandSelector.Span.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                WorkbookMutationAction.GroupRows.expanded())));
    assertEquals(
        List.of(ColumnBandSelector.Insertion.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new WorkbookMutationAction.InsertColumns())));
    assertEquals(
        List.of(ColumnBandSelector.Span.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                WorkbookMutationAction.GroupColumns.expanded())));
    assertEquals(
        List.of(DrawingObjectSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new DrawingMutationAction.DeleteDrawingObject())));
    assertEquals(
        List.of(PivotTableSelector.ByNameOnSheet.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.DeletePivotTable())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth())));
    assertEquals(
        List.of(
            CellSelector.AllUsedInSheet.class,
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new SheetIntrospectionQuery.GetComments())));
    assertEquals(
        List.of(RangeSelector.RectangularWindow.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionSurfaceQuery.GetSheetSchema())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new SheetIntrospectionQuery.GetConditionalFormatting())));
    assertEquals(
        List.of(DrawingObjectSelector.AllOnSheet.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects())));
    assertEquals(
        List.of(DrawingObjectSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload())));
    assertEquals(
        List.of(PivotTableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzePivotTableHealth())));
    assertEquals(
        List.of(TableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeTableHealth())));
  }

  @Test
  void coversEveryRemainingGroupedActionAndQueryCaseAndThreeWayUnionWording() {
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new DrawingMutationAction.SetPicture(
                    new dev.erst.gridgrind.contract.dto.PictureInput(
                        "Logo",
                        new dev.erst.gridgrind.contract.dto.PictureDataInput(
                            dev.erst.gridgrind.excel.foundation.ExcelPictureFormat.PNG,
                            binary("AQID")),
                        new dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell(
                            new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(0, 0, 0, 0),
                            new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(1, 1, 0, 0),
                            dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior
                                .MOVE_AND_RESIZE),
                        java.util.Optional.empty())))));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new DrawingMutationAction.SetSignatureLine(
                    new dev.erst.gridgrind.contract.dto.SignatureLineInput(
                        "Signer",
                        new dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell(
                            new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(0, 0, 0, 0),
                            new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(1, 1, 0, 0),
                            dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior
                                .MOVE_AND_RESIZE),
                        false,
                        java.util.Optional.of("Review before signing."),
                        java.util.Optional.of("Ada Lovelace"),
                        java.util.Optional.of("Finance"),
                        java.util.Optional.of("ada@example.com"),
                        java.util.Optional.empty(),
                        java.util.Optional.empty(),
                        java.util.Optional.of(
                            new dev.erst.gridgrind.contract.dto.PictureDataInput(
                                dev.erst.gridgrind.excel.foundation.ExcelPictureFormat.PNG,
                                binary("AQID"))))))));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookMutationAction.SetWorkbookProtection(
                    new dev.erst.gridgrind.contract.dto.WorkbookProtectionInput(
                        true,
                        false,
                        false,
                        java.util.Optional.empty(),
                        java.util.Optional.empty())))));
    assertEquals(
        List.of(RangeSelector.ByRange.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new WorkbookMutationAction.MergeCells())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.ClearDataValidations())));
    assertEquals(
        List.of(ColumnBandSelector.Span.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookMutationAction.SetColumnWidth(8.43d))));
    assertEquals(
        List.of(RowBandSelector.Span.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookMutationAction.SetRowHeight(15.0d))));
    assertEquals(
        List.of(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new CellMutationAction.ClearHyperlink())));
    assertEquals(
        List.of(DrawingObjectSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new DrawingMutationAction.SetDrawingObjectAnchor(
                    new dev.erst.gridgrind.contract.dto.DrawingAnchorInput.TwoCell(
                        new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(0, 0, 0, 0),
                        new dev.erst.gridgrind.contract.dto.DrawingMarkerInput(1, 1, 0, 0),
                        dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior
                            .MOVE_AND_RESIZE)))));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.SetTable(
                    dev.erst.gridgrind.contract.dto.TableInput.withDefaultMetadata(
                        "BudgetTable",
                        "Budget",
                        "A1:B2",
                        false,
                        new dev.erst.gridgrind.contract.dto.TableStyleInput.None())))));
    assertEquals(
        List.of(PivotTableSelector.ByNameOnSheet.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.SetPivotTable(
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
                                dev.erst.gridgrind.excel.foundation
                                    .ExcelPivotDataConsolidateFunction.SUM,
                                "Total Amount",
                                Optional.empty())))))));
    assertEquals(
        List.of(
            NamedRangeSelector.ByName.class,
            NamedRangeSelector.WorkbookScope.class,
            NamedRangeSelector.SheetScope.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.SetNamedRange(
                    "BudgetTotal",
                    new dev.erst.gridgrind.contract.dto.NamedRangeScope.Workbook(),
                    dev.erst.gridgrind.contract.dto.NamedRangeTarget.range("Budget", "A1")))));

    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())));
    assertEquals(
        List.of(NamedRangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookIntrospectionQuery.GetNamedRanges())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new SheetIntrospectionQuery.GetSheetLayout())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionSurfaceQuery.GetFormulaSurface())));
    assertEquals(
        List.of(
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new SheetIntrospectionQuery.GetCells())));
    assertEquals(
        List.of(RangeSelector.RectangularWindow.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new SheetIntrospectionQuery.GetWindow())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new SheetIntrospectionQuery.GetDataValidations())));
    assertEquals(
        List.of(PivotTableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookAssetIntrospectionQuery.GetPivotTables())));
    assertEquals(
        List.of(TableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookAssetIntrospectionQuery.GetTables())));

    IllegalArgumentException threeWayUnionFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.requireCompatible(
                    new WorkbookSelector.Current(),
                    new StructuredMutationAction.DeleteNamedRange()));
    assertEquals(
        "DELETE_NAMED_RANGE requires target type NAMED_RANGE_BY_NAME, NAMED_RANGE_WORKBOOK_SCOPE"
            + " or NAMED_RANGE_SHEET_SCOPE but got WORKBOOK_CURRENT",
        threeWayUnionFailure.getMessage());

    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookMutationAction.AutoSizeColumns())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.ClearAutofilter())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new CellMutationAction.AppendRow(
                    new dev.erst.gridgrind.contract.dto.CellRowInput.Typed(
                        List.of(new CellInput.Text(text("A"))))))));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookMutationAction.ClearWorkbookProtection())));
    assertEquals(
        List.of(RangeSelector.ByRange.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new CellMutationAction.SetRange(
                    new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                        List.of(List.of(new CellInput.Text(text("A")))))))));
    assertEquals(
        List.of(RangeSelector.ByRange.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new WorkbookMutationAction.UnmergeCells())));
    assertEquals(
        List.of(RangeSelector.ByRange.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new CellMutationAction.ClearRange())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.ClearConditionalFormatting())));
    assertEquals(
        List.of(ColumnBandSelector.Span.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new WorkbookMutationAction.DeleteColumns())));
    assertEquals(
        List.of(RowBandSelector.Span.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new WorkbookMutationAction.DeleteRows())));
    assertEquals(
        List.of(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(new CellMutationAction.ClearComment())));
    assertEquals(
        List.of(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new CellMutationAction.SetHyperlink(
                    new dev.erst.gridgrind.contract.dto.HyperlinkTarget.Url(
                        "https://example.com")))));
    assertEquals(
        List.of(DrawingObjectSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new DrawingMutationAction.DeleteDrawingObject())));
    assertEquals(
        List.of(TableSelector.ByNameOnSheet.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new StructuredMutationAction.DeleteTable())));
    assertEquals(
        List.of(PivotTableSelector.ByNameOnSheet.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.DeletePivotTable())));
    assertEquals(
        List.of(
            NamedRangeSelector.ByName.class,
            NamedRangeSelector.WorkbookScope.class,
            NamedRangeSelector.SheetScope.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new StructuredMutationAction.DeleteNamedRange())));

    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookIntrospectionQuery.GetWorkbookProtection())));
    assertEquals(
        List.of(WorkbookSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new WorkbookIntrospectionQuery.GetPackageSecurity())));
    assertEquals(
        List.of(NamedRangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionSurfaceQuery.GetNamedRangeSurface())));
    assertEquals(
        List.of(NamedRangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeNamedRangeHealth())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new SheetIntrospectionQuery.GetAutofilters())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new SheetIntrospectionQuery.GetMergedRegions())));
    assertEquals(
        List.of(SheetSelector.ByName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new SheetIntrospectionQuery.GetPrintLayout())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeHyperlinkHealth())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeDataValidationHealth())));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeAutofilterHealth())));
    assertEquals(
        List.of(
            CellSelector.AllUsedInSheet.class,
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new SheetIntrospectionQuery.GetHyperlinks())));
    assertEquals(
        List.of(
            CellSelector.AllUsedInSheet.class,
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(new SheetIntrospectionQuery.GetComments())));
    assertEquals(
        List.of(RangeSelector.RectangularWindow.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionSurfaceQuery.GetSheetSchema())));
    assertEquals(
        List.of(RangeSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new SheetIntrospectionQuery.GetConditionalFormatting())));
    assertEquals(
        List.of(PivotTableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzePivotTableHealth())));
    assertEquals(
        List.of(TableSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new InspectionAnalysisQuery.AnalyzeTableHealth())));
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
        new CellMutationAction.SetCell(new CellInput.NumberValue(125.0)),
        WorkbookStepValidation.requireCompatible(
            tableCell, new CellMutationAction.SetCell(new CellInput.NumberValue(125.0))));
    assertEquals(
        new SheetIntrospectionQuery.GetCells(),
        WorkbookStepValidation.requireCompatible(
            tableCell, new SheetIntrospectionQuery.GetCells()));
    assertEquals(
        new CellAssertion.DisplayValue("125"),
        WorkbookStepValidation.requireCompatible(tableCell, new CellAssertion.DisplayValue("125")));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }

  private static BinarySourceInput binary(String value) {
    return BinarySourceInput.inlineBase64(value);
  }
}
