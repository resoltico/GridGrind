package dev.erst.gridgrind.core.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for package-private helper seams inside the authoring module. */
class AuthoringInternalsCoverageTest {
  @Test
  void plannedStepsPreserveExplicitIdsAndRejectInvalidInputs() {
    SheetSelector.ByName sheetSelector = new SheetSelector.ByName(new SheetName("Budget"));
    CellSelector.ByAddress cellSelector =
        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1"));
    WorkbookMutationAction.EnsureSheet ensureSheet = new WorkbookMutationAction.EnsureSheet();
    SheetIntrospectionQuery.GetCells cellQuery = new SheetIntrospectionQuery.GetCells();
    CellAssertion.Value present = new CellAssertion.Value(ExpectedValues.blank());

    PlannedMutation unnamedMutation = new PlannedMutation(sheetSelector, ensureSheet);
    PlannedInspection unnamedInspection = new PlannedInspection(cellSelector, cellQuery);
    PlannedAssertion unnamedAssertion = new PlannedAssertion(cellSelector, present);

    MutationStep generatedMutation = unnamedMutation.toStep(new StepId("mutation-001"));
    InspectionStep generatedInspection = unnamedInspection.toStep(new StepId("inspection-001"));
    AssertionStep generatedAssertion = unnamedAssertion.toStep(new StepId("assertion-001"));
    assertEquals("mutation-001", generatedMutation.stepId().value());
    assertEquals("inspection-001", generatedInspection.stepId().value());
    assertEquals("assertion-001", generatedAssertion.stepId().value());

    assertEquals(
        "named-mutation",
        unnamedMutation
            .named(new StepId("named-mutation"))
            .toStep(new StepId("ignored"))
            .stepId()
            .value());
    assertEquals(
        "named-inspection",
        unnamedInspection
            .named(new StepId("named-inspection"))
            .toStep(new StepId("ignored"))
            .stepId()
            .value());
    assertEquals(
        "named-assertion",
        unnamedAssertion
            .named(new StepId("named-assertion"))
            .toStep(new StepId("ignored"))
            .stepId()
            .value());

    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlannedMutation(Optional.of(new StepId(" ")), sheetSelector, ensureSheet))
            .getMessage());
    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlannedInspection(Optional.of(new StepId(" ")), cellSelector, cellQuery))
            .getMessage());
    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlannedAssertion(Optional.of(new StepId(" ")), cellSelector, present))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedMutation(Optional.of(new StepId("id")), null, ensureSheet))
            .getMessage());
    assertEquals(
        "action must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedMutation(Optional.of(new StepId("id")), sheetSelector, null))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedInspection(Optional.of(new StepId("id")), null, cellQuery))
            .getMessage());
    assertEquals(
        "query must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedInspection(Optional.of(new StepId("id")), cellSelector, null))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedAssertion(Optional.of(new StepId("id")), null, present))
            .getMessage());
    assertEquals(
        "assertion must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new PlannedAssertion(Optional.of(new StepId("id")), cellSelector, null))
            .getMessage());
  }

  @Test
  void checksCoverAllSupportedAssertionHelpers() {
    InspectionAnalysisQuery.AnalyzeFormulaHealth analysisQuery =
        InspectionAnalysisQueries.formulaHealth();

    assertInstanceOf(PresenceAssertion.NamedRangePresent.class, Checks.namedRangePresent());
    assertInstanceOf(PresenceAssertion.NamedRangeAbsent.class, Checks.namedRangeAbsent());
    assertInstanceOf(PresenceAssertion.TablePresent.class, Checks.tablePresent());
    assertInstanceOf(PresenceAssertion.TableAbsent.class, Checks.tableAbsent());
    assertInstanceOf(PresenceAssertion.PivotTablePresent.class, Checks.pivotTablePresent());
    assertInstanceOf(PresenceAssertion.PivotTableAbsent.class, Checks.pivotTableAbsent());
    assertInstanceOf(PresenceAssertion.ChartPresent.class, Checks.chartPresent());
    assertInstanceOf(PresenceAssertion.ChartAbsent.class, Checks.chartAbsent());
    assertInstanceOf(CellAssertion.Value.class, Checks.cellValue(ExpectedValues.text("Owner")));
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
        WorkbookIntrospectionQuery.GetWorkbookSummary.class, WorkbookQueries.workbookSummary());
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetPackageSecurity.class, WorkbookQueries.packageSecurity());
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetWorkbookProtection.class,
        WorkbookQueries.workbookProtection());
    assertInstanceOf(
        WorkbookIntrospectionQuery.GetNamedRanges.class, WorkbookQueries.namedRanges());
    assertInstanceOf(SheetIntrospectionQuery.GetSheetSummary.class, SheetQueries.sheetSummary());
    assertInstanceOf(SheetIntrospectionQuery.GetCells.class, SheetQueries.cells());
    assertInstanceOf(SheetIntrospectionQuery.GetWindow.class, SheetQueries.window());
    assertInstanceOf(SheetIntrospectionQuery.GetMergedRegions.class, SheetQueries.mergedRegions());
    assertInstanceOf(SheetIntrospectionQuery.GetHyperlinks.class, SheetQueries.hyperlinks());
    assertInstanceOf(SheetIntrospectionQuery.GetComments.class, SheetQueries.comments());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetDrawingObjects.class,
        WorkbookAssetQueries.drawingObjects());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetCharts.class, WorkbookAssetQueries.charts());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetPivotTables.class, WorkbookAssetQueries.pivotTables());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload.class,
        WorkbookAssetQueries.drawingObjectPayload());
    assertInstanceOf(SheetIntrospectionQuery.GetSheetLayout.class, SheetQueries.sheetLayout());
    assertInstanceOf(SheetIntrospectionQuery.GetPrintLayout.class, SheetQueries.printLayout());
    assertInstanceOf(
        SheetIntrospectionQuery.GetDataValidations.class, SheetQueries.dataValidations());
    assertInstanceOf(
        SheetIntrospectionQuery.GetConditionalFormatting.class,
        SheetQueries.conditionalFormatting());
    assertInstanceOf(SheetIntrospectionQuery.GetAutofilters.class, SheetQueries.autofilters());
    assertInstanceOf(
        WorkbookAssetIntrospectionQuery.GetTables.class, WorkbookAssetQueries.tables());
    assertInstanceOf(
        InspectionSurfaceQuery.GetFormulaSurface.class, InspectionSurfaceQueries.formulaSurface());
    assertInstanceOf(
        InspectionSurfaceQuery.GetSheetSchema.class, InspectionSurfaceQueries.sheetSchema());
    assertInstanceOf(
        InspectionSurfaceQuery.GetNamedRangeSurface.class,
        InspectionSurfaceQueries.namedRangeSurface());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeFormulaHealth.class,
        InspectionAnalysisQueries.formulaHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeDataValidationHealth.class,
        InspectionAnalysisQueries.dataValidationHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeConditionalFormattingHealth.class,
        InspectionAnalysisQueries.conditionalFormattingHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeAutofilterHealth.class,
        InspectionAnalysisQueries.autofilterHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeTableHealth.class, InspectionAnalysisQueries.tableHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzePivotTableHealth.class,
        InspectionAnalysisQueries.pivotTableHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeHyperlinkHealth.class,
        InspectionAnalysisQueries.hyperlinkHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeNamedRangeHealth.class,
        InspectionAnalysisQueries.namedRangeHealth());
    assertInstanceOf(
        InspectionAnalysisQuery.AnalyzeWorkbookFindings.class,
        InspectionAnalysisQueries.workbookFindings());
  }
}
