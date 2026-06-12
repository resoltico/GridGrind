package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for the ordered step-based workbook plan contract. */
class WorkbookPlanTest {
  @Test
  void standardFactoryBuildsExplicitDefaultSectionsAndRejectsNullSteps() {
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());

    assertEquals(GridGrindProtocolVersion.current(), plan.protocolVersion());
    assertInstanceOf(WorkbookPlan.WorkbookPersistence.None.class, plan.persistence());
    assertEquals(java.util.Optional.empty(), plan.planId());
    assertTrue(plan.execution().isDefault());
    assertTrue(plan.executionMode().isDefault());
    assertEquals(ExecutionJournalLevel.SUMMARY, plan.journalLevel());
    assertTrue(plan.formulaEnvironment().isEmpty());
    assertEquals(List.of(), plan.steps());
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookPlan(
                GridGrindProtocolVersion.current(),
                Optional.empty(),
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                ExecutionPolicyInput.defaults(),
                FormulaEnvironmentInput.empty(),
                null));
  }

  @Test
  void copiesStepsAndRejectsDuplicateStepIds() {
    WorkbookStep authoredStep =
        new MutationStep(
            "ensure-budget",
            new SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.EnsureSheet());
    List<WorkbookStep> steps = new ArrayList<>();
    steps.add(authoredStep);
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            steps);

    steps.clear();

    assertEquals(1, plan.steps().size());
    assertEquals("ensure-budget", plan.steps().getFirst().stepId());
    assertThrows(UnsupportedOperationException.class, () -> plan.steps().add(authoredStep));

    IllegalArgumentException duplicateStepFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookPlan.standard(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionPolicyInput.defaults(),
                    FormulaEnvironmentInput.empty(),
                    List.of(
                        new MutationStep(
                            "duplicate",
                            new SheetSelector.ByName("Budget"),
                            new WorkbookMutationAction.EnsureSheet()),
                        new InspectionStep(
                            "duplicate",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookSummary()))));

    assertEquals(
        "steps must not contain duplicate stepId values: duplicate",
        duplicateStepFailure.getMessage());
    assertInstanceOf(InvalidRequestException.class, duplicateStepFailure);
    assertEquals(
        Optional.of("steps[1].stepId"),
        ((InvalidRequestException) duplicateStepFailure).jsonPath());
  }

  @Test
  void existingAndSaveAsWorkbookPathsMustPointToXlsxFiles() {
    assertEquals("budget.xlsx", new WorkbookPlan.WorkbookSource.ExistingFile("budget.xlsx").path());
    assertEquals("budget.xlsx", new WorkbookPlan.WorkbookPersistence.SaveAs("budget.xlsx").path());

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookPlan.WorkbookSource.ExistingFile("budget.xlsm"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookPlan.WorkbookPersistence.SaveAs("budget.xls"));
  }

  @Test
  void supportsExecutionModeAndFormulaEnvironmentConstructors() {
    FormulaEnvironmentInput formulaEnvironment =
        new FormulaEnvironmentInput(
            List.of(new FormulaExternalWorkbookInput("rates.xlsx", "tmp/rates.xlsx")),
            FormulaMissingWorkbookPolicy.USE_CACHED_VALUE,
            List.of());
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("budget.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(ExecutionModeInput.eventRead()),
            formulaEnvironment,
            List.of(
                new MutationStep(
                    "set-cell",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(new CellInput.Text(text("Owner"))))));

    assertInstanceOf(ExecutionModeInput.EventRead.class, plan.executionMode());
    assertInstanceOf(ExecutionModeInput.EventRead.class, plan.effectiveExecutionMode());
    assertEquals(formulaEnvironment, plan.formulaEnvironment());
    assertEquals("set-cell", plan.steps().getFirst().stepId());
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookPlan(
                GridGrindProtocolVersion.current(),
                Optional.empty(),
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                null,
                formulaEnvironment,
                List.of()));
  }

  @Test
  void supportsExplicitPlanIdAndExecutionJournalPolicy() {
    WorkbookPlan plan =
        WorkbookPlan.identified(
            GridGrindProtocolVersion.current(),
            "budget-audit",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.modeAndJournal(
                ExecutionModeInput.fullXssf(),
                new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
            FormulaEnvironmentInput.empty(),
            List.of());

    assertEquals("budget-audit", plan.planId().orElseThrow());
    assertEquals(ExecutionJournalLevel.VERBOSE, plan.journalLevel());
    assertEquals(ExecutionJournalLevel.VERBOSE, plan.execution().journal().level());
    assertInstanceOf(ExecutionModeInput.FullXssf.class, plan.effectiveExecutionMode());
    assertEquals(
        "planId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    WorkbookPlan.identified(
                        GridGrindProtocolVersion.current(),
                        " ",
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        ExecutionPolicyInput.defaults(),
                        FormulaEnvironmentInput.empty(),
                        List.of()))
            .getMessage());
  }

  @Test
  void supportsExecutionPolicyConstructorAndDefaultEffectiveExecution() {
    ExecutionPolicyInput executionPolicy =
        ExecutionPolicyInput.modeAndJournal(
            ExecutionModeInput.eventRead(),
            new ExecutionJournalInput(ExecutionJournalLevel.SUMMARY));
    WorkbookPlan explicitPlan =
        new WorkbookPlan(
            GridGrindProtocolVersion.current(),
            Optional.empty(),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy,
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan defaultPlan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());

    assertEquals(executionPolicy, explicitPlan.execution());
    assertEquals(ExecutionJournalLevel.SUMMARY, explicitPlan.journalLevel());
    assertInstanceOf(ExecutionModeInput.EventRead.class, explicitPlan.executionMode());
    assertEquals(
        ExecutionJournalLevel.SUMMARY, defaultPlan.effectiveExecution().effectiveJournalLevel());
    assertInstanceOf(ExecutionModeInput.FullXssf.class, defaultPlan.effectiveExecutionMode());
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookPlan(
                GridGrindProtocolVersion.current(),
                Optional.empty(),
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                null,
                FormulaEnvironmentInput.empty(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookPlan(
                GridGrindProtocolVersion.current(),
                Optional.empty(),
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                executionPolicy,
                null,
                List.of()));
  }

  @Test
  void separatesMutationAssertionAndInspectionViewsAndValidatesWorkbookPaths() {
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("budget.xlsx"),
            new WorkbookPlan.WorkbookPersistence.SaveAs("report.xlsx"),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-cell",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(new CellInput.Text(text("Owner")))),
                new AssertionStep(
                    "assert-cell",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellAssertion.CellValue(
                        new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner"))),
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));

    assertEquals(1, plan.stepPartition().mutations().size());
    assertEquals(1, plan.stepPartition().assertions().size());
    assertEquals(1, plan.stepPartition().inspections().size());
    assertEquals("set-cell", plan.stepPartition().mutations().getFirst().stepId());
    assertEquals("assert-cell", plan.stepPartition().assertions().getFirst().stepId());
    assertEquals("summary", plan.stepPartition().inspections().getFirst().stepId());
    InvalidRequestException nullStepFailure =
        assertThrows(
            InvalidRequestException.class,
            () ->
                WorkbookPlan.standard(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionPolicyInput.defaults(),
                    FormulaEnvironmentInput.empty(),
                    java.util.Arrays.asList(
                        new MutationStep(
                            "ok",
                            new WorkbookSelector.Current(),
                            new WorkbookMutationAction.ClearWorkbookProtection()),
                        null)));
    assertEquals("steps must not contain nulls", nullStepFailure.getMessage());
    assertEquals(Optional.of("steps[1]"), nullStepFailure.jsonPath());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookPlan.WorkbookSource.ExistingFile(" "))
            .getMessage()
            .startsWith("path "));
  }

  @Test
  void rejectsStepCountExceedingMaximum() {
    List<WorkbookStep> manySteps = new ArrayList<>();
    for (int i = 0; i <= WorkbookPlan.MAX_STEPS; i++) {
      manySteps.add(
          new MutationStep(
              "step-" + i,
              new SheetSelector.ByName("Sheet1"),
              new WorkbookMutationAction.EnsureSheet()));
    }

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookPlan.standard(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionPolicyInput.defaults(),
                    FormulaEnvironmentInput.empty(),
                    manySteps));
    assertTrue(failure.getMessage().contains(String.valueOf(WorkbookPlan.MAX_STEPS)));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }
}
