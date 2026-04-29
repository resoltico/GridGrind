package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.selector.CellSelector;
import java.util.List;
import java.util.Optional;
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

    CalculationPolicyInput defaultPolicy = CalculationPolicyInput.defaults();
    CalculationPolicyInput markedPolicy =
        new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true);
    CalculationPolicyInput targetedPolicy = new CalculationPolicyInput(evaluateTargets);

    assertTrue(defaultPolicy.isDefault());
    assertEquals("DO_NOT_CALCULATE", defaultPolicy.effectiveStrategy().strategyType());
    assertFalse(markedPolicy.isDefault());
    assertTrue(markedPolicy.markRecalculateOnOpen());
    assertEquals(evaluateTargets, targetedPolicy.effectiveStrategy());

    ExecutionPolicyInput defaultExecution =
        new ExecutionPolicyInput(
            ExecutionModeInput.defaults(),
            ExecutionJournalInput.defaults(),
            CalculationPolicyInput.defaults());
    ExecutionPolicyInput customExecution = ExecutionPolicyInput.calculation(markedPolicy);

    assertTrue(defaultExecution.isDefault());
    assertEquals(ExecutionModeInput.defaults(), defaultExecution.mode());
    assertEquals(ExecutionJournalInput.defaults(), defaultExecution.journal());
    assertEquals(CalculationPolicyInput.defaults(), defaultExecution.calculation());
    assertFalse(customExecution.isDefault());
    assertEquals(markedPolicy, customExecution.effectiveCalculation());
    assertTrue(
        GridGrindContractText.calculationPolicyInputSummary().contains("markRecalculateOnOpen"));
    assertTrue(
        GridGrindContractText.calculationStrategyInputSummary().contains("EVALUATE_TARGETS"));
  }

  @Test
  void executionPolicyFiltersAndNormalizedExecutionDefaultsBehavePrecisely() {
    ExecutionPolicyInput.DefaultFilter policyFilter = new ExecutionPolicyInput.DefaultFilter();
    ExecutionPolicyInput.ExecutionModeDefaultFilter modeFilter =
        new ExecutionPolicyInput.ExecutionModeDefaultFilter();
    ExecutionPolicyInput.ExecutionJournalDefaultFilter journalFilter =
        new ExecutionPolicyInput.ExecutionJournalDefaultFilter();
    ExecutionPolicyInput.CalculationPolicyDefaultFilter calculationFilter =
        new ExecutionPolicyInput.CalculationPolicyDefaultFilter();
    ExecutionPolicyInput defaultPolicy = ExecutionPolicyInput.defaults();
    ExecutionPolicyInput customPolicy =
        new ExecutionPolicyInput(
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.EVENT_READ,
                ExecutionModeInput.WriteMode.STREAMING_WRITE),
            new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE),
            new CalculationPolicyInput(new CalculationStrategyInput.EvaluateAll(), true));
    ExecutionModeInput defaultMode = ExecutionModeInput.defaults();
    ExecutionModeInput customMode =
        new ExecutionModeInput(
            ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.STREAMING_WRITE);
    ExecutionJournalInput defaultJournal = ExecutionJournalInput.defaults();
    ExecutionJournalInput customJournal = new ExecutionJournalInput(ExecutionJournalLevel.SUMMARY);
    CalculationPolicyInput defaultCalculation = CalculationPolicyInput.defaults();
    CalculationPolicyInput customCalculation =
        new CalculationPolicyInput(new CalculationStrategyInput.EvaluateAll(), true);

    assertTrue(filterMatches(policyFilter, null));
    assertTrue(filterMatches(policyFilter, defaultPolicy));
    assertFalse(filterMatches(policyFilter, customPolicy));
    assertFalse(filterMatches(policyFilter, "policy"));
    assertEquals(0, policyFilter.hashCode());

    assertTrue(filterMatches(modeFilter, null));
    assertTrue(filterMatches(modeFilter, defaultMode));
    assertFalse(filterMatches(modeFilter, customMode));
    assertFalse(filterMatches(modeFilter, "mode"));
    assertEquals(0, modeFilter.hashCode());

    assertTrue(filterMatches(journalFilter, null));
    assertTrue(filterMatches(journalFilter, defaultJournal));
    assertFalse(filterMatches(journalFilter, customJournal));
    assertFalse(filterMatches(journalFilter, "journal"));
    assertEquals(0, journalFilter.hashCode());

    assertTrue(filterMatches(calculationFilter, null));
    assertTrue(filterMatches(calculationFilter, defaultCalculation));
    assertFalse(filterMatches(calculationFilter, customCalculation));
    assertFalse(filterMatches(calculationFilter, "calculation"));
    assertEquals(0, calculationFilter.hashCode());

    assertTrue(defaultMode.isDefault());
    assertFalse(customMode.isDefault());
    assertTrue(defaultJournal.isDefault());
    assertFalse(customJournal.isDefault());
    assertEquals(ExecutionJournalLevel.NORMAL, ExecutionJournalInput.effectiveLevel(null));
  }

  @Test
  void validatesCalculationReportsAndJournalCalculationEnvelope() {
    CalculationReport.FormulaCapability evaluable =
        new CalculationReport.FormulaCapability(
            new CellSelector.QualifiedAddress("Budget", "B1"),
            "A1*2",
            FormulaCapabilityKind.EVALUABLE_NOW,
            Optional.empty(),
            Optional.empty());
    CalculationReport.FormulaCapability blocking =
        new CalculationReport.FormulaCapability(
            new CellSelector.QualifiedAddress("Budget", "C1"),
            "APP.TITLE()",
            FormulaCapabilityKind.UNEVALUABLE_NOW,
            Optional.of(GridGrindProblemCode.UNSUPPORTED_FORMULA),
            Optional.of("Unsupported formula function APP.TITLE at Budget!C1: APP.TITLE()"));
    CalculationReport.Preflight preflight =
        new CalculationReport.Preflight(
            CalculationReport.Scope.TARGETS,
            2,
            new CalculationReport.Summary(1, 1, 0),
            List.of(evaluable, blocking));
    CalculationReport.Execution execution =
        new CalculationReport.Execution(CalculationExecutionStatus.SUCCEEDED, 1, false, true);
    CalculationReport report =
        new CalculationReport(
            new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true),
            preflight,
            execution);

    assertTrue(CalculationReport.notRequested().policy().isDefault());
    assertTrue(
        CalculationReport.create(null, java.util.Optional.empty(), execution).policy().isDefault());
    assertEquals(2, report.preflight().orElseThrow().checkedFormulaCount());
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
                Optional.of(GridGrindProblemCode.INVALID_FORMULA),
                Optional.of("unexpected")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.FormulaCapability(
                new CellSelector.QualifiedAddress("Budget", "B1"),
                "A1*2",
                FormulaCapabilityKind.UNPARSEABLE_BY_POI,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.FormulaCapability(
                new CellSelector.QualifiedAddress("Budget", "B1"),
                "A1*2",
                FormulaCapabilityKind.EVALUABLE_NOW,
                Optional.empty(),
                Optional.of("unexpected")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.Execution(
                CalculationExecutionStatus.SUCCEEDED, -1, false, false));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.Execution(
                CalculationExecutionStatus.FAILED, 0, false, false, Optional.of(" ")));
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
  void executeCalculationContextMergesTypedLocationsWithoutCrossShapePadding() {
    ProblemContext.ExecuteCalculation base =
        new ProblemContext.ExecuteCalculation.Preflight(
            ProblemContext.RequestShape.known("EXISTING", "SAVE_AS"),
            ProblemContext.ProblemLocation.unknown());
    ProblemContext.ExecuteCalculation enriched =
        base.withLocation(
            ProblemContext.ProblemLocation.formulaCell("Budget", "B1", "APP.TITLE()"));
    ProblemContext.ExecuteCalculation preserved =
        new ProblemContext.ExecuteCalculation.Execution(
                ProblemContext.RequestShape.known("EXISTING", "SAVE_AS"),
                ProblemContext.ProblemLocation.formulaCell("Ops", "C4", "SUM(A1:A3)"))
            .withLocation(
                ProblemContext.ProblemLocation.formulaCell("Ignored", "Ignored", "Ignored"));

    assertEquals("CALCULATION_PREFLIGHT", enriched.stage());
    assertEquals(java.util.Optional.of("Budget"), enriched.sheetName());
    assertEquals(java.util.Optional.of("B1"), enriched.address());
    assertEquals(java.util.Optional.of("APP.TITLE()"), enriched.formula());
    assertEquals(java.util.Optional.of("Ops"), preserved.sheetName());
    assertEquals(java.util.Optional.of("C4"), preserved.address());
    assertEquals(java.util.Optional.of("SUM(A1:A3)"), preserved.formula());
  }

  private static boolean filterMatches(Object filter, Object candidate) {
    return filter.equals(candidate);
  }
}
