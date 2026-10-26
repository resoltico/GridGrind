package dev.erst.gridgrind.core.model.formula;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.catalog.GridGrindContractText;
import dev.erst.gridgrind.core.plan.AssertionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces;
import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct validation coverage for Phase 6 calculation contract types. */
class CalculationContractTypesTest {
  @Test
  void validatesCalculationStrategiesAndExecutionPolicyDefaults() {
    CalculationStrategySpec.DoNotCalculate doNotCalculate =
        new CalculationStrategySpec.DoNotCalculate();
    CalculationStrategySpec.DeferredCalculation deferredCalculation =
        new CalculationStrategySpec.DeferredCalculation();
    CalculationStrategySpec.EvaluateAll evaluateAll = new CalculationStrategySpec.EvaluateAll();
    CalculationStrategySpec.ClearCachesOnly clearCachesOnly =
        new CalculationStrategySpec.ClearCachesOnly();
    CalculationStrategySpec.EvaluateTargets evaluateTargets =
        new CalculationStrategySpec.EvaluateTargets(
            List.of(new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1"))));

    assertTrue(doNotCalculate.isDefault());
    assertEquals("DO_NOT_CALCULATE", doNotCalculate.strategyType());
    assertFalse(deferredCalculation.isDefault());
    assertEquals("DEFERRED_CALCULATION", deferredCalculation.strategyType());
    assertFalse(evaluateAll.isDefault());
    assertEquals("EVALUATE_ALL", evaluateAll.strategyType());
    assertEquals("CLEAR_CACHES_ONLY", clearCachesOnly.strategyType());
    assertEquals("EVALUATE_TARGETS", evaluateTargets.strategyType());
    assertThrows(
        NullPointerException.class, () -> new CalculationStrategySpec.EvaluateTargets(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CalculationStrategySpec.EvaluateTargets(List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new CalculationStrategySpec.EvaluateTargets(List.of((FormulaCellTarget) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationStrategySpec.EvaluateTargets(
                List.of(
                    new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
                    new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")))));

    CalculationPolicySpec defaultPolicy = CalculationPolicySpec.defaults();
    CalculationPolicySpec markedPolicy =
        new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), true);
    CalculationPolicySpec targetedPolicy = CalculationPolicySpec.strategy(evaluateTargets);

    assertTrue(defaultPolicy.isDefault());
    assertEquals("DO_NOT_CALCULATE", defaultPolicy.effectiveStrategy().strategyType());
    assertFalse(markedPolicy.isDefault());
    assertTrue(markedPolicy.markRecalculateOnOpen());
    assertEquals(evaluateTargets, targetedPolicy.effectiveStrategy());

    ExecutionPolicySpec defaultExecution =
        new ExecutionPolicySpec(
            ExecutionModeSpec.defaults(),
            ExecutionJournalSpec.defaults(),
            CalculationPolicySpec.defaults(),
            AssertionModeSpec.defaults());
    ExecutionPolicySpec customExecution = ExecutionPolicySpec.calculation(markedPolicy);

    assertTrue(defaultExecution.isDefault());
    assertEquals(ExecutionModeSpec.defaults(), defaultExecution.mode());
    assertEquals(ExecutionJournalSpec.defaults(), defaultExecution.journal());
    assertEquals(CalculationPolicySpec.defaults(), defaultExecution.calculation());
    assertEquals(AssertionModeSpec.FAIL_FAST, defaultExecution.assertionMode());
    assertFalse(customExecution.isDefault());
    assertEquals(markedPolicy, customExecution.effectiveCalculation());
    assertTrue(
        GridGrindContractText.calculationPolicyInputSummary().contains("markRecalculateOnOpen"));
    assertTrue(
        GridGrindContractText.calculationStrategyInputSummary().contains("EVALUATE_TARGETS"));
  }

  @Test
  void executionPolicyFactoriesAndNullRejectionBehavePrecisely() {
    ExecutionPolicySpec defaultPolicy = ExecutionPolicySpec.defaults();
    ExecutionPolicySpec customPolicy =
        new ExecutionPolicySpec(
            ExecutionModeSpec.eventRead(),
            new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE),
            new CalculationPolicySpec(new CalculationStrategySpec.EvaluateAll(), true),
            AssertionModeSpec.defaults());
    ExecutionModeSpec defaultMode = ExecutionModeSpec.defaults();
    ExecutionModeSpec customMode = ExecutionModeSpec.eventRead();
    ExecutionJournalSpec defaultJournal = ExecutionJournalSpec.defaults();
    ExecutionJournalSpec customJournal = new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE);
    CalculationPolicySpec defaultCalculation = CalculationPolicySpec.defaults();
    CalculationPolicySpec customCalculation =
        new CalculationPolicySpec(new CalculationStrategySpec.EvaluateAll(), true);

    assertEquals(defaultPolicy, ExecutionPolicySpec.mode(defaultMode));
    assertEquals(defaultPolicy, ExecutionPolicySpec.journal(defaultJournal));
    assertEquals(defaultPolicy, ExecutionPolicySpec.calculation(defaultCalculation));
    assertEquals(
        new ExecutionPolicySpec(
            customMode, defaultJournal, defaultCalculation, AssertionModeSpec.defaults()),
        ExecutionPolicySpec.mode(customMode));
    assertEquals(
        new ExecutionPolicySpec(
            defaultMode, customJournal, defaultCalculation, AssertionModeSpec.defaults()),
        ExecutionPolicySpec.journal(customJournal));
    assertEquals(
        new ExecutionPolicySpec(
            customMode, customJournal, defaultCalculation, AssertionModeSpec.defaults()),
        ExecutionPolicySpec.modeAndJournal(customMode, customJournal));
    assertEquals(
        new ExecutionPolicySpec(
            customMode, defaultJournal, customCalculation, AssertionModeSpec.defaults()),
        ExecutionPolicySpec.modeAndCalculation(customMode, customCalculation));
    assertEquals(
        new ExecutionPolicySpec(
            defaultMode, defaultJournal, defaultCalculation, AssertionModeSpec.COLLECT),
        ExecutionPolicySpec.assertionMode(AssertionModeSpec.COLLECT));
    assertFalse(ExecutionPolicySpec.assertionMode(AssertionModeSpec.COLLECT).isDefault());
    assertFalse(customPolicy.isDefault());
    assertTrue(defaultMode.isDefault());
    assertFalse(customMode.isDefault());
    assertEquals("EVENT_READ", customMode.modeType());
    assertEquals("STREAMING_WRITE", ExecutionModeSpec.streamingWrite().modeType());
    assertTrue(defaultJournal.isDefault());
    assertFalse(customJournal.isDefault());
    assertEquals(
        ExecutionJournalLevel.SUMMARY, ExecutionJournalSpec.effectiveLevel(defaultJournal));
    assertThrows(NullPointerException.class, () -> ExecutionJournalSpec.effectiveLevel(null));
  }

  @Test
  void validatesCalculationReportsAndJournalCalculationEnvelope() {
    CalculationReport.FormulaCapability evaluable =
        new CalculationReport.FormulaCapability(
            new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
            "A1*2",
            FormulaCapabilityKind.EVALUABLE_NOW,
            Optional.empty(),
            Optional.empty());
    CalculationReport.FormulaCapability blocking =
        new CalculationReport.FormulaCapability(
            new FormulaCellTarget(new SheetName("Budget"), new CellAddress("C1")),
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
            new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), true),
            preflight,
            execution);

    assertTrue(CalculationReport.notRequested().policy().isDefault());
    assertTrue(
        new CalculationReport(
                CalculationPolicySpec.defaults(), java.util.Optional.empty(), execution)
            .policy()
            .isDefault());
    assertEquals(
        "policy must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new CalculationReport(null, java.util.Optional.empty(), execution))
            .getMessage());
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
                new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
                "A1*2",
                FormulaCapabilityKind.EVALUABLE_NOW,
                Optional.of(GridGrindProblemCode.INVALID_FORMULA),
                Optional.of("unexpected")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.FormulaCapability(
                new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
                "A1*2",
                FormulaCapabilityKind.UNPARSEABLE_BY_POI,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CalculationReport.FormulaCapability(
                new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
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
            ProblemContextRequestSurfaces.RequestShape.known("EXISTING", "SAVE_AS"),
            ProblemContextWorkbookSurfaces.ProblemLocation.unknown());
    ProblemContext.ExecuteCalculation enriched =
        base.withLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                new SheetLabel("Budget"), "B1", "APP.TITLE()"));
    ProblemContext.ExecuteCalculation preserved =
        new ProblemContext.ExecuteCalculation.Execution(
                ProblemContextRequestSurfaces.RequestShape.known("EXISTING", "SAVE_AS"),
                ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                    new SheetLabel("Ops"), "C4", "SUM(A1:A3)"))
            .withLocation(
                ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                    new SheetLabel("Ignored"), "Ignored", "Ignored"));

    assertEquals("CALCULATION_PREFLIGHT", enriched.type());
    assertEquals(
        java.util.Optional.of("Budget"),
        enriched.sheetName().map(dev.erst.gridgrind.core.vocabulary.WireString::value));
    assertEquals(java.util.Optional.of("B1"), enriched.address());
    assertEquals(java.util.Optional.of("APP.TITLE()"), enriched.formula());
    assertEquals(
        java.util.Optional.of("Ops"),
        preserved.sheetName().map(dev.erst.gridgrind.core.vocabulary.WireString::value));
    assertEquals(java.util.Optional.of("C4"), preserved.address());
    assertEquals(java.util.Optional.of("SUM(A1:A3)"), preserved.formula());
  }
}
