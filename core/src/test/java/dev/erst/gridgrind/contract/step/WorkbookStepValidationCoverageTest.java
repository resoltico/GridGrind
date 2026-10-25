package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Additional branch coverage for selector-family inference across assertion workflows. */
class WorkbookStepValidationCoverageTest {
  @Test
  void resolvesCompositeAndNegatedAssertionTargetFamilies() {
    CompositeAssertion.AllOf allOf =
        new CompositeAssertion.AllOf(
            List.of(new PresenceAssertion.TablePresent(), new PresenceAssertion.TableAbsent()));
    CompositeAssertion.AnyOf anyOf =
        new CompositeAssertion.AnyOf(
            List.of(new PresenceAssertion.TablePresent(), new PresenceAssertion.TableAbsent()));
    CompositeAssertion.Not not =
        new CompositeAssertion.Not(
            new CellAssertion.CellValue(
                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner")));

    assertEquals(
        List.of(TableSelector.class),
        List.of(WorkbookOperationContracts.targetSelectorsFor(allOf)));
    assertEquals(
        List.of(TableSelector.class),
        List.of(WorkbookOperationContracts.targetSelectorsFor(anyOf)));
    assertEquals(
        List.of(
            CellSelector.ByAddress.class,
            CellSelector.ByAddresses.class,
            TableCellSelector.ByColumnName.class),
        List.of(WorkbookOperationContracts.targetSelectorsFor(not)));
  }

  @Test
  void resolvesAnalysisTargetFamiliesForAllAssertionQueryForms() {
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookOperationContracts.targetSelectorsFor(
                new AnalysisAssertion.AnalysisMaxSeverity(
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                    AnalysisSeverity.WARNING))));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookOperationContracts.targetSelectorsFor(
                new AnalysisAssertion.AnalysisFindingPresent(
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    Optional.of(AnalysisSeverity.ERROR),
                    Optional.empty()))));
    assertEquals(
        List.of(SheetSelector.class),
        List.of(
            WorkbookOperationContracts.targetSelectorsFor(
                new AnalysisAssertion.AnalysisFindingAbsent(
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                    AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                    Optional.empty(),
                    Optional.empty()))));
  }

  @Test
  void rejectsEmptyCompositeIntersections() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookOperationContracts.targetSelectorsFor(
                    new CompositeAssertion.AnyOf(
                        List.of(
                            new PresenceAssertion.TablePresent(),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.Text(
                                    "Owner"))))));

    assertEquals(
        "ANY_OF requires nested assertions with compatible target families", failure.getMessage());
  }
}
