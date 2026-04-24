package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Additional branch coverage for selector-family inference across assertion workflows. */
class WorkbookStepValidationCoverageTest {
  @Test
  void resolvesCompositeAndNegatedAssertionTargetFamilies() {
    Assertion.AllOf allOf =
        new Assertion.AllOf(List.of(new Assertion.TablePresent(), new Assertion.TableAbsent()));
    Assertion.AnyOf anyOf =
        new Assertion.AnyOf(List.of(new Assertion.TablePresent(), new Assertion.TableAbsent()));
    Assertion.Not not =
        new Assertion.Not(new Assertion.CellValue(new ExpectedCellValue.Text("Owner")));

    assertEquals(
        List.of(TableSelector.class), List.of(WorkbookStepValidation.allowedTargetTypes(allOf)));
    assertEquals(
        List.of(TableSelector.class), List.of(WorkbookStepValidation.allowedTargetTypes(anyOf)));
    assertEquals(
        List.of(
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(WorkbookStepValidation.allowedTargetTypes(not)));
  }

  @Test
  void resolvesAnalysisTargetFamiliesForAllAssertionQueryForms() {
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new Assertion.AnalysisMaxSeverity(
                    new InspectionQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING))));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new Assertion.AnalysisFindingPresent(
                    new InspectionQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    AnalysisSeverity.ERROR,
                    null))));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookStepValidation.allowedTargetTypes(
                new Assertion.AnalysisFindingAbsent(
                    new InspectionQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                    null,
                    null))));
  }

  @Test
  void rejectsEmptyCompositeIntersections() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookStepValidation.allowedTargetTypes(
                    new Assertion.AnyOf(
                        List.of(
                            new Assertion.TablePresent(),
                            new Assertion.CellValue(new ExpectedCellValue.Text("Owner"))))));

    assertEquals(
        "ANY_OF requires nested assertions with compatible target families", failure.getMessage());
  }

  @Test
  void rejectsEmptyNestedAssertionCollectionsAtCommonTargetInferenceBoundary() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> WorkbookStepValidation.commonTargetTypes(List.of(), "ALL_OF"));

    assertEquals(
        "ALL_OF requires nested assertions with compatible target families", failure.getMessage());
  }
}
