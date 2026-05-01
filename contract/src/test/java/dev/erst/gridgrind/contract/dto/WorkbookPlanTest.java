package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.query.InspectionQuery;
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
    assertEquals(ExecutionJournalLevel.NORMAL, plan.journalLevel());
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
                            new InspectionQuery.GetWorkbookSummary()))));

    assertEquals(
        "steps must not contain duplicate stepId values: duplicate",
        duplicateStepFailure.getMessage());
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
            ExecutionPolicyInput.mode(
                new ExecutionModeInput(
                    ExecutionModeInput.ReadMode.EVENT_READ,
                    ExecutionModeInput.WriteMode.FULL_XSSF)),
            formulaEnvironment,
            List.of(
                new MutationStep(
                    "set-cell",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(new CellInput.Text(text("Owner"))))));

    assertEquals(ExecutionModeInput.ReadMode.EVENT_READ, plan.executionMode().readMode());
    assertEquals(ExecutionModeInput.ReadMode.EVENT_READ, plan.effectiveExecutionMode().readMode());
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
                new ExecutionModeInput(
                    ExecutionModeInput.ReadMode.FULL_XSSF, ExecutionModeInput.WriteMode.FULL_XSSF),
                new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
            FormulaEnvironmentInput.empty(),
            List.of());

    assertEquals("budget-audit", plan.planId().orElseThrow());
    assertEquals(ExecutionJournalLevel.VERBOSE, plan.journalLevel());
    assertEquals(ExecutionJournalLevel.VERBOSE, plan.execution().journal().level());
    assertEquals(ExecutionModeInput.ReadMode.FULL_XSSF, plan.effectiveExecutionMode().readMode());
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
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF),
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
    assertEquals(ExecutionModeInput.ReadMode.EVENT_READ, explicitPlan.executionMode().readMode());
    assertEquals(
        ExecutionJournalLevel.NORMAL, defaultPlan.effectiveExecution().effectiveJournalLevel());
    assertEquals(
        ExecutionModeInput.ReadMode.FULL_XSSF, defaultPlan.effectiveExecutionMode().readMode());
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
                    new Assertion.CellValue(new ExpectedCellValue.Text("Owner"))),
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())));

    assertEquals(1, plan.stepPartition().mutations().size());
    assertEquals(1, plan.stepPartition().assertions().size());
    assertEquals(1, plan.stepPartition().inspections().size());
    assertEquals("set-cell", plan.stepPartition().mutations().getFirst().stepId());
    assertEquals("assert-cell", plan.stepPartition().assertions().getFirst().stepId());
    assertEquals("summary", plan.stepPartition().inspections().getFirst().stepId());
    assertEquals(
        "steps must not contain nulls",
        assertThrows(
                NullPointerException.class,
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
                            null)))
            .getMessage());
    assertTrue(
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookPlan.WorkbookSource.ExistingFile(" "))
            .getMessage()
            .startsWith("path "));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }
}
