package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.selector.CellSelector;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct validation coverage for Phase 6 calculation contract types. */
class CalculationContractTypesTest {
  @Test
  void validatesCalculationStrategiesAndExecutionPolicyDefaults() {
    CalculationStrategyInput.DoNotCalculate doNotCalculate =
        new CalculationStrategyInput.DoNotCalculate();
    CalculationStrategyInput.EvaluateAll evaluateAll = new CalculationStrategyInput.EvaluateAll();
    CalculationStrategyInput.ClearCachesOnly clearCachesOnly =
        new CalculationStrategyInput.ClearCachesOnly();
    CalculationStrategyInput.EvaluateTargets evaluateTargets =
        new CalculationStrategyInput.EvaluateTargets(
            List.of(new CellSelector.QualifiedAddress("Budget", "B1")));

    assertTrue(doNotCalculate.isDefault());
    assertEquals("DO_NOT_CALCULATE", doNotCalculate.strategyType());
    assertFalse(evaluateAll.isDefault());
    assertEquals("EVALUATE_ALL", evaluateAll.strategyType());
    assertEquals("CLEAR_CACHES_ONLY", clearCachesOnly.strategyType());
    assertEquals("EVALUATE_TARGETS", evaluateTargets.strategyType());
    assertThrows(
        NullPointerException.class, () -> new CalculationStrategyInput.EvaluateTargets(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CalculationStrategyInput.EvaluateTargets(List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new CalculationStrategyInput.EvaluateTargets(
                List.of((CellSelector.QualifiedAddress) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationStrategyInput.EvaluateTargets(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "B1"),
                    new CellSelector.QualifiedAddress("Budget", "B1"))));

    CalculationPolicyInput defaultPolicy = new CalculationPolicyInput(null, false);
    CalculationPolicyInput markedPolicy = CalculationPolicyInput.create(null, true);
    CalculationPolicyInput targetedPolicy = new CalculationPolicyInput(evaluateTargets);

    assertTrue(defaultPolicy.isDefault());
    assertEquals("DO_NOT_CALCULATE", defaultPolicy.effectiveStrategy().strategyType());
    assertFalse(markedPolicy.isDefault());
    assertTrue(markedPolicy.markRecalculateOnOpen());
    assertEquals(evaluateTargets, targetedPolicy.effectiveStrategy());

    ExecutionPolicyInput defaultExecution =
        new ExecutionPolicyInput(
            new ExecutionModeInput(null, null),
            new ExecutionJournalInput(null),
            new CalculationPolicyInput(null, false));
    ExecutionPolicyInput customExecution = new ExecutionPolicyInput(null, null, markedPolicy);

    assertTrue(defaultExecution.isDefault());
    assertNull(defaultExecution.mode());
    assertNull(defaultExecution.journal());
    assertNull(defaultExecution.calculation());
    assertFalse(customExecution.isDefault());
    assertEquals(markedPolicy, customExecution.effectiveCalculation());
    assertTrue(
        GridGrindContractText.calculationPolicyInputSummary().contains("markRecalculateOnOpen"));
    assertTrue(
        GridGrindContractText.calculationStrategyInputSummary().contains("EVALUATE_TARGETS"));
  }

  @Test
  void validatesCalculationReportsAndJournalCalculationEnvelope() {
    CalculationReport.FormulaCapability evaluable =
        new CalculationReport.FormulaCapability(
            new CellSelector.QualifiedAddress("Budget", "B1"),
            "A1*2",
            FormulaCapabilityKind.EVALUABLE_NOW,
            null,
            null);
    CalculationReport.FormulaCapability blocking =
        new CalculationReport.FormulaCapability(
            new CellSelector.QualifiedAddress("Budget", "C1"),
            "APP.TITLE()",
            FormulaCapabilityKind.UNEVALUABLE_NOW,
            GridGrindProblemCode.UNSUPPORTED_FORMULA,
            "Unsupported formula function APP.TITLE at Budget!C1: APP.TITLE()");
    CalculationReport.Preflight preflight =
        new CalculationReport.Preflight(
            CalculationReport.Scope.TARGETS,
            2,
            new CalculationReport.Summary(1, 1, 0),
            List.of(evaluable, blocking));
    CalculationReport.Execution execution =
        new CalculationReport.Execution(CalculationExecutionStatus.SUCCEEDED, 1, false, true, null);
    CalculationReport report =
        new CalculationReport(new CalculationPolicyInput(null, true), preflight, execution);

    assertTrue(CalculationReport.notRequested().policy().isDefault());
    assertTrue(new CalculationReport(null, null, execution).policy().isDefault());
    assertEquals(2, report.preflight().checkedFormulaCount());
    assertTrue(report.execution().markRecalculateOnOpenApplied());
    assertEquals(
        "scope must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new CalculationReport.Preflight(
                        null, 0, new CalculationReport.Summary(0, 0, 0), List.of()))
            .getMessage());
    assertEquals(
        "checkedFormulaCount must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CalculationReport.Preflight(
                        CalculationReport.Scope.WORKBOOK,
                        -1,
                        new CalculationReport.Summary(0, 0, 0),
                        List.of()))
            .getMessage());
    assertEquals(
        "summary must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new CalculationReport.Preflight(
                        CalculationReport.Scope.WORKBOOK, 0, null, List.of()))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.Preflight(
                CalculationReport.Scope.WORKBOOK,
                2,
                new CalculationReport.Summary(1, 0, 0),
                List.of(evaluable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.Preflight(
                CalculationReport.Scope.WORKBOOK,
                1,
                new CalculationReport.Summary(1, 1, 0),
                List.of(evaluable)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.FormulaCapability(
                new CellSelector.QualifiedAddress("Budget", "B1"),
                "A1*2",
                FormulaCapabilityKind.EVALUABLE_NOW,
                GridGrindProblemCode.INVALID_FORMULA,
                "unexpected"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.FormulaCapability(
                new CellSelector.QualifiedAddress("Budget", "B1"),
                "A1*2",
                FormulaCapabilityKind.UNPARSEABLE_BY_POI,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.Execution(
                CalculationExecutionStatus.SUCCEEDED, -1, false, false, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.Execution(
                CalculationExecutionStatus.FAILED, 0, false, false, " "));
    assertThrows(IllegalArgumentException.class, () -> new CalculationReport.Summary(-1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CalculationReport.Summary(0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new CalculationReport.Summary(0, 0, -1));

    ExecutionJournal.Phase phase = ExecutionJournal.Phase.notStarted();
    ExecutionJournal.Calculation calculation = new ExecutionJournal.Calculation(phase, phase);
    assertEquals(phase, calculation.preflight());
    assertEquals(
        "execution must not be null",
        assertThrows(
                IllegalArgumentException.class, () -> new ExecutionJournal.Calculation(phase, null))
            .getMessage());
  }

  @Test
  void executeCalculationContextMergesExceptionFactsWithoutOverwritingExistingValues() {
    GridGrindResponse.ProblemContext.ExecuteCalculation base =
        new GridGrindResponse.ProblemContext.ExecuteCalculation.Preflight(
            "EXISTING", "SAVE_AS", null, null, null);
    GridGrindResponse.ProblemContext.ExecuteCalculation enriched =
        base.withExceptionData("Budget", "B1", "APP.TITLE()");
    GridGrindResponse.ProblemContext.ExecuteCalculation preserved =
        new GridGrindResponse.ProblemContext.ExecuteCalculation.Execution(
                "EXISTING", "SAVE_AS", "Ops", "C4", "SUM(A1:A3)")
            .withExceptionData("Ignored", "Ignored", "Ignored");

    assertEquals("CALCULATION_PREFLIGHT", enriched.stage());
    assertEquals("Budget", enriched.sheetName());
    assertEquals("B1", enriched.address());
    assertEquals("APP.TITLE()", enriched.formula());
    assertEquals("Ops", preserved.sheetName());
    assertEquals("C4", preserved.address());
    assertEquals("SUM(A1:A3)", preserved.formula());
  }
}
