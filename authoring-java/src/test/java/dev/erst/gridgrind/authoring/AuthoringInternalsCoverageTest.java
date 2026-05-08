package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for package-private helper seams inside the authoring module. */
class AuthoringInternalsCoverageTest {
  @Test
  void plannedStepsPreserveExplicitIdsAndRejectInvalidInputs() {
    SheetSelector.ByName sheetSelector = new SheetSelector.ByName("Budget");
    CellSelector.ByAddress cellSelector = new CellSelector.ByAddress("Budget", "A1");
    WorkbookMutationAction.EnsureSheet ensureSheet = new WorkbookMutationAction.EnsureSheet();
    SheetIntrospectionQuery.GetCells cellQuery = new SheetIntrospectionQuery.GetCells();
    CellAssertion.CellValue present =
        new CellAssertion.CellValue(Values.toExpectedCellValue(Values.expectedBlank()));

    PlannedMutation unnamedMutation = new PlannedMutation(sheetSelector, ensureSheet);
    PlannedInspection unnamedInspection = new PlannedInspection(cellSelector, cellQuery);
    PlannedAssertion unnamedAssertion = new PlannedAssertion(cellSelector, present);

    MutationStep generatedMutation = unnamedMutation.toStep("mutation-001");
    InspectionStep generatedInspection = unnamedInspection.toStep("inspection-001");
    AssertionStep generatedAssertion = unnamedAssertion.toStep("assertion-001");
    assertEquals("mutation-001", generatedMutation.stepId());
    assertEquals("inspection-001", generatedInspection.stepId());
    assertEquals("assertion-001", generatedAssertion.stepId());

    assertEquals(
        "named-mutation", unnamedMutation.named("named-mutation").toStep("ignored").stepId());
    assertEquals(
        "named-inspection", unnamedInspection.named("named-inspection").toStep("ignored").stepId());
    assertEquals(
        "named-assertion", unnamedAssertion.named("named-assertion").toStep("ignored").stepId());

    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlannedMutation(Optional.of(" "), sheetSelector, ensureSheet))
            .getMessage());
    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlannedInspection(Optional.of(" "), cellSelector, cellQuery))
            .getMessage());
    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlannedAssertion(Optional.of(" "), cellSelector, present))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedMutation(Optional.of("id"), null, ensureSheet))
            .getMessage());
    assertEquals(
        "action must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedMutation(Optional.of("id"), sheetSelector, null))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedInspection(Optional.of("id"), null, cellQuery))
            .getMessage());
    assertEquals(
        "query must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedInspection(Optional.of("id"), cellSelector, null))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedAssertion(Optional.of("id"), null, present))
            .getMessage());
    assertEquals(
        "assertion must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedAssertion(Optional.of("id"), cellSelector, null))
            .getMessage());
  }

  @Test
  void checksCoverAllSupportedAssertionHelpers() {
    InspectionAnalysisQuery.AnalyzeFormulaHealth analysisQuery = Queries.formulaHealth();

    assertInstanceOf(PresenceAssertion.NamedRangePresent.class, Checks.namedRangePresent());
    assertInstanceOf(PresenceAssertion.NamedRangeAbsent.class, Checks.namedRangeAbsent());
    assertInstanceOf(PresenceAssertion.TablePresent.class, Checks.tablePresent());
    assertInstanceOf(PresenceAssertion.TableAbsent.class, Checks.tableAbsent());
    assertInstanceOf(PresenceAssertion.PivotTablePresent.class, Checks.pivotTablePresent());
    assertInstanceOf(PresenceAssertion.PivotTableAbsent.class, Checks.pivotTableAbsent());
    assertInstanceOf(PresenceAssertion.ChartPresent.class, Checks.chartPresent());
    assertInstanceOf(PresenceAssertion.ChartAbsent.class, Checks.chartAbsent());
    assertInstanceOf(CellAssertion.CellValue.class, Checks.cellValue(Values.expectedText("Owner")));
    assertInstanceOf(CellAssertion.DisplayValue.class, Checks.displayValue("Owner"));
    assertInstanceOf(CellAssertion.FormulaText.class, Checks.formulaText("SUM(A1:A2)"));
    assertInstanceOf(
        AnalysisAssertion.AnalysisMaxSeverity.class,
        Checks.analysisMaxSeverity(analysisQuery, AnalysisSeverity.WARNING));
    assertInstanceOf(
        AnalysisAssertion.AnalysisFindingPresent.class,
        Checks.analysisFindingPresent(
            analysisQuery,
            AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
            AnalysisSeverity.ERROR,
            "detail"));
    assertInstanceOf(
        AnalysisAssertion.AnalysisFindingAbsent.class,
        Checks.analysisFindingAbsent(
            analysisQuery,
            AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
            AnalysisSeverity.ERROR,
            "detail"));
    CompositeAssertion.AllOf allOf = Checks.allOf(Checks.chartPresent(), Checks.chartAbsent());
    CompositeAssertion.AnyOf anyOf = Checks.anyOf(Checks.tablePresent(), Checks.tableAbsent());
    CompositeAssertion.Not not = Checks.not(Checks.namedRangePresent());
    assertEquals(2, allOf.assertions().size());
    assertEquals(2, anyOf.assertions().size());
    assertInstanceOf(PresenceAssertion.NamedRangePresent.class, not.assertion());
    assertEquals(
        "assertion must not be null",
        assertThrows(NullPointerException.class, () -> Checks.not(null)).getMessage());
  }

  @Test
  void queriesCoverAllSupportedInspectionHelpers() {
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetWorkbookSummary.class, Queries.workbookSummary());
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetPackageSecurity.class, Queries.packageSecurity());
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetWorkbookProtection.class, Queries.workbookProtection());
    assertInstanceOf(WorkbookIntrospectionQuery.GetNamedRanges.class, Queries.namedRanges());
    assertInstanceOf(SheetIntrospectionQuery.GetSheetSummary.class, Queries.sheetSummary());
    assertInstanceOf(SheetIntrospectionQuery.GetCells.class, Queries.cells());
    assertInstanceOf(SheetIntrospectionQuery.GetWindow.class, Queries.window());
    assertInstanceOf(SheetIntrospectionQuery.GetMergedRegions.class, Queries.mergedRegions());
    assertInstanceOf(SheetIntrospectionQuery.GetHyperlinks.class, Queries.hyperlinks());
    assertInstanceOf(SheetIntrospectionQuery.GetComments.class, Queries.comments());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetDrawingObjects.class, Queries.drawingObjects());
    assertInstanceOf(WorkbookAssetIntrospectionQuery.GetCharts.class, Queries.charts());
    assertInstanceOf(WorkbookAssetIntrospectionQuery.GetPivotTables.class, Queries.pivotTables());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload.class,
        Queries.drawingObjectPayload());
    assertInstanceOf(SheetIntrospectionQuery.GetSheetLayout.class, Queries.sheetLayout());
    assertInstanceOf(SheetIntrospectionQuery.GetPrintLayout.class, Queries.printLayout());
    assertInstanceOf(SheetIntrospectionQuery.GetDataValidations.class, Queries.dataValidations());
    assertInstanceOf(
        SheetIntrospectionQuery.GetConditionalFormatting.class, Queries.conditionalFormatting());
    assertInstanceOf(SheetIntrospectionQuery.GetAutofilters.class, Queries.autofilters());
    assertInstanceOf(WorkbookAssetIntrospectionQuery.GetTables.class, Queries.tables());
    assertInstanceOf(InspectionSurfaceQuery.GetFormulaSurface.class, Queries.formulaSurface());
    assertInstanceOf(InspectionSurfaceQuery.GetSheetSchema.class, Queries.sheetSchema());
    assertInstanceOf(
        InspectionSurfaceQuery.GetNamedRangeSurface.class, Queries.namedRangeSurface());
    assertInstanceOf(InspectionAnalysisQuery.AnalyzeFormulaHealth.class, Queries.formulaHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeDataValidationHealth.class, Queries.dataValidationHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth.class,
        Queries.conditionalFormattingHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeAutofilterHealth.class, Queries.autofilterHealth());
    assertInstanceOf(InspectionAnalysisQuery.AnalyzeTableHealth.class, Queries.tableHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzePivotTableHealth.class, Queries.pivotTableHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeHyperlinkHealth.class, Queries.hyperlinkHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeNamedRangeHealth.class, Queries.namedRangeHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeWorkbookFindings.class, Queries.workbookFindings());
  }
}
