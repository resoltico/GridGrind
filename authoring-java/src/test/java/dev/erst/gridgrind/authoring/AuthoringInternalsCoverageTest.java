package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import org.junit.jupiter.api.Test;

/** Focused coverage for package-private helper seams inside the authoring module. */
class AuthoringInternalsCoverageTest {
  @Test
  void plannedStepsPreserveExplicitIdsAndRejectInvalidInputs() {
    SheetSelector.ByName sheetSelector = new SheetSelector.ByName("Budget");
    CellSelector.ByAddress cellSelector = new CellSelector.ByAddress("Budget", "A1");
    MutationAction.EnsureSheet ensureSheet = new MutationAction.EnsureSheet();
    InspectionQuery.GetCells cellQuery = new InspectionQuery.GetCells();
    Assertion.CellValue present =
        new Assertion.CellValue(Values.toExpectedCellValue(Values.expectedBlank()));

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
                () -> new PlannedMutation(" ", sheetSelector, ensureSheet))
            .getMessage());
    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlannedInspection(" ", cellSelector, cellQuery))
            .getMessage());
    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new PlannedAssertion(" ", cellSelector, present))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(NullPointerException.class, () -> new PlannedMutation("id", null, ensureSheet))
            .getMessage());
    assertEquals(
        "action must not be null",
        assertThrows(
                NullPointerException.class, () -> new PlannedMutation("id", sheetSelector, null))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(NullPointerException.class, () -> new PlannedInspection("id", null, cellQuery))
            .getMessage());
    assertEquals(
        "query must not be null",
        assertThrows(
                NullPointerException.class, () -> new PlannedInspection("id", cellSelector, null))
            .getMessage());
    assertEquals(
        "target must not be null",
        assertThrows(NullPointerException.class, () -> new PlannedAssertion("id", null, present))
            .getMessage());
    assertEquals(
        "assertion must not be null",
        assertThrows(
                NullPointerException.class, () -> new PlannedAssertion("id", cellSelector, null))
            .getMessage());
  }

  @Test
  void checksCoverAllSupportedAssertionHelpers() {
    InspectionQuery.AnalyzeFormulaHealth analysisQuery = Queries.formulaHealth();

    assertInstanceOf(Assertion.Present.class, Checks.present());
    assertInstanceOf(Assertion.Absent.class, Checks.absent());
    assertInstanceOf(Assertion.CellValue.class, Checks.cellValue(Values.expectedText("Owner")));
    assertInstanceOf(Assertion.DisplayValue.class, Checks.displayValue("Owner"));
    assertInstanceOf(Assertion.FormulaText.class, Checks.formulaText("SUM(A1:A2)"));
    assertInstanceOf(
        Assertion.AnalysisMaxSeverity.class,
        Checks.analysisMaxSeverity(analysisQuery, AnalysisSeverity.WARNING));
    assertInstanceOf(
        Assertion.AnalysisFindingPresent.class,
        Checks.analysisFindingPresent(
            analysisQuery,
            AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
            AnalysisSeverity.ERROR,
            "detail"));
    assertInstanceOf(
        Assertion.AnalysisFindingAbsent.class,
        Checks.analysisFindingAbsent(
            analysisQuery,
            AnalysisFindingCode.PIVOT_TABLE_UNSUPPORTED_DETAIL,
            AnalysisSeverity.ERROR,
            "detail"));
    Assertion.AllOf allOf = Checks.allOf(Checks.present(), Checks.absent());
    Assertion.AnyOf anyOf = Checks.anyOf(Checks.present(), Checks.absent());
    Assertion.Not not = Checks.not(Checks.present());
    assertEquals(2, allOf.assertions().size());
    assertEquals(2, anyOf.assertions().size());
    assertInstanceOf(Assertion.Present.class, not.assertion());
    assertEquals(
        "assertion must not be null",
        assertThrows(NullPointerException.class, () -> Checks.not(null)).getMessage());
  }

  @Test
  void queriesCoverAllSupportedInspectionHelpers() {
    assertInstanceOf(InspectionQuery.GetWorkbookSummary.class, Queries.workbookSummary());
    assertInstanceOf(InspectionQuery.GetPackageSecurity.class, Queries.packageSecurity());
    assertInstanceOf(InspectionQuery.GetWorkbookProtection.class, Queries.workbookProtection());
    assertInstanceOf(InspectionQuery.GetNamedRanges.class, Queries.namedRanges());
    assertInstanceOf(InspectionQuery.GetSheetSummary.class, Queries.sheetSummary());
    assertInstanceOf(InspectionQuery.GetCells.class, Queries.cells());
    assertInstanceOf(InspectionQuery.GetWindow.class, Queries.window());
    assertInstanceOf(InspectionQuery.GetMergedRegions.class, Queries.mergedRegions());
    assertInstanceOf(InspectionQuery.GetHyperlinks.class, Queries.hyperlinks());
    assertInstanceOf(InspectionQuery.GetComments.class, Queries.comments());
    assertInstanceOf(InspectionQuery.GetDrawingObjects.class, Queries.drawingObjects());
    assertInstanceOf(InspectionQuery.GetCharts.class, Queries.charts());
    assertInstanceOf(InspectionQuery.GetPivotTables.class, Queries.pivotTables());
    assertInstanceOf(InspectionQuery.GetDrawingObjectPayload.class, Queries.drawingObjectPayload());
    assertInstanceOf(InspectionQuery.GetSheetLayout.class, Queries.sheetLayout());
    assertInstanceOf(InspectionQuery.GetPrintLayout.class, Queries.printLayout());
    assertInstanceOf(InspectionQuery.GetDataValidations.class, Queries.dataValidations());
    assertInstanceOf(
        InspectionQuery.GetConditionalFormatting.class, Queries.conditionalFormatting());
    assertInstanceOf(InspectionQuery.GetAutofilters.class, Queries.autofilters());
    assertInstanceOf(InspectionQuery.GetTables.class, Queries.tables());
    assertInstanceOf(InspectionQuery.GetFormulaSurface.class, Queries.formulaSurface());
    assertInstanceOf(InspectionQuery.GetSheetSchema.class, Queries.sheetSchema());
    assertInstanceOf(InspectionQuery.GetNamedRangeSurface.class, Queries.namedRangeSurface());
    assertInstanceOf(InspectionQuery.AnalyzeFormulaHealth.class, Queries.formulaHealth());
    assertInstanceOf(
        InspectionQuery.AnalyzeDataValidationHealth.class, Queries.dataValidationHealth());
    assertInstanceOf(
        InspectionQuery.AnalyzeConditionalFormattingHealth.class,
        Queries.conditionalFormattingHealth());
    assertInstanceOf(InspectionQuery.AnalyzeAutofilterHealth.class, Queries.autofilterHealth());
    assertInstanceOf(InspectionQuery.AnalyzeTableHealth.class, Queries.tableHealth());
    assertInstanceOf(InspectionQuery.AnalyzePivotTableHealth.class, Queries.pivotTableHealth());
    assertInstanceOf(InspectionQuery.AnalyzeHyperlinkHealth.class, Queries.hyperlinkHealth());
    assertInstanceOf(InspectionQuery.AnalyzeNamedRangeHealth.class, Queries.namedRangeHealth());
    assertInstanceOf(InspectionQuery.AnalyzeWorkbookFindings.class, Queries.workbookFindings());
  }
}
