package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
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
import org.junit.jupiter.api.Test;

/** Tests for the ordered step-based workbook plan contract. */
class WorkbookPlanTest {
  @Test
  void defaultsProtocolVersionPersistenceAndOptionalSections() {
    WorkbookPlan plan = new WorkbookPlan(new WorkbookPlan.WorkbookSource.New(), null, List.of());
    WorkbookPlan nullStepsPlan =
        new WorkbookPlan(new WorkbookPlan.WorkbookSource.New(), null, (List<WorkbookStep>) null);

    assertEquals(GridGrindProtocolVersion.current(), plan.protocolVersion());
    assertInstanceOf(WorkbookPlan.WorkbookPersistence.None.class, plan.persistence());
    assertNull(plan.planId());
    assertNull(plan.execution());
    assertNull(plan.executionMode());
    assertEquals(ExecutionJournalLevel.NORMAL, plan.journalLevel());
    assertNull(plan.formulaEnvironment());
    assertEquals(List.of(), plan.steps());
    assertEquals(List.of(), nullStepsPlan.steps());
  }

  @Test
  void copiesStepsAndRejectsDuplicateStepIds() {
    List<WorkbookStep> steps = new ArrayList<>();
    steps.add(
        new MutationStep(
            "ensure-budget", new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()));
    WorkbookPlan plan =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            steps);

    steps.clear();

    assertEquals(1, plan.steps().size());
    assertEquals("ensure-budget", plan.steps().getFirst().stepId());

    IllegalArgumentException duplicateStepFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new WorkbookPlan(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(
                        new MutationStep(
                            "duplicate",
                            new SheetSelector.ByName("Budget"),
                            new MutationAction.EnsureSheet()),
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
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.ExistingFile("budget.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF),
            formulaEnvironment,
            List.of(
                new MutationStep(
                    "set-cell",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(new CellInput.Text(text("Owner"))))));

    assertEquals(ExecutionModeInput.ReadMode.EVENT_READ, plan.executionMode().readMode());
    assertEquals(ExecutionModeInput.ReadMode.EVENT_READ, plan.effectiveExecutionMode().readMode());
    assertEquals(formulaEnvironment, plan.formulaEnvironment());
    assertEquals("set-cell", plan.steps().getFirst().stepId());
    assertNull(
        new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                null,
                new FormulaEnvironmentInput(
                    List.of(), FormulaMissingWorkbookPolicy.ERROR, List.of()),
                List.of())
            .formulaEnvironment());
  }

  @Test
  void supportsExplicitPlanIdAndExecutionJournalPolicy() {
    WorkbookPlan plan =
        new WorkbookPlan(
            GridGrindProtocolVersion.current(),
            "budget-audit",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionPolicyInput(
                new ExecutionModeInput(
                    ExecutionModeInput.ReadMode.FULL_XSSF, ExecutionModeInput.WriteMode.FULL_XSSF),
                new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
            null,
            List.of());

    assertEquals("budget-audit", plan.planId());
    assertEquals(ExecutionJournalLevel.VERBOSE, plan.journalLevel());
    assertEquals(ExecutionJournalLevel.VERBOSE, plan.execution().journal().level());
    assertEquals(ExecutionModeInput.ReadMode.FULL_XSSF, plan.effectiveExecutionMode().readMode());
    assertEquals(
        "planId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookPlan(
                        GridGrindProtocolVersion.current(),
                        " ",
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        null,
                        null,
                        List.of()))
            .getMessage());
  }

  @Test
  void supportsExecutionPolicyConstructorAndDefaultEffectiveExecution() {
    ExecutionPolicyInput executionPolicy =
        new ExecutionPolicyInput(
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF),
            new ExecutionJournalInput(ExecutionJournalLevel.SUMMARY));
    WorkbookPlan explicitPlan =
        new WorkbookPlan(
            GridGrindProtocolVersion.current(),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy,
            null,
            List.of());
    WorkbookPlan defaultPlan =
        new WorkbookPlan(
            GridGrindProtocolVersion.current(),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            (ExecutionPolicyInput) null,
            null,
            List.of());

    assertEquals(executionPolicy, explicitPlan.execution());
    assertEquals(ExecutionJournalLevel.SUMMARY, explicitPlan.journalLevel());
    assertEquals(ExecutionModeInput.ReadMode.EVENT_READ, explicitPlan.executionMode().readMode());
    assertEquals(
        ExecutionJournalLevel.NORMAL, defaultPlan.effectiveExecution().effectiveJournalLevel());
    assertEquals(
        ExecutionModeInput.ReadMode.FULL_XSSF, defaultPlan.effectiveExecutionMode().readMode());
  }

  @Test
  void separatesMutationAssertionAndInspectionViewsAndValidatesWorkbookPaths() {
    WorkbookPlan plan =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.ExistingFile("budget.xlsx"),
            new WorkbookPlan.WorkbookPersistence.SaveAs("report.xlsx"),
            List.of(
                new MutationStep(
                    "set-cell",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new MutationAction.SetCell(new CellInput.Text(text("Owner")))),
                new AssertionStep(
                    "assert-cell",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new Assertion.CellValue(new ExpectedCellValue.Text("Owner"))),
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())));

    assertEquals(1, plan.mutationSteps().size());
    assertEquals(1, plan.assertionSteps().size());
    assertEquals(1, plan.inspectionSteps().size());
    assertEquals("set-cell", plan.mutationSteps().getFirst().stepId());
    assertEquals("assert-cell", plan.assertionSteps().getFirst().stepId());
    assertEquals("summary", plan.inspectionSteps().getFirst().stepId());
    assertEquals(
        "steps must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new WorkbookPlan(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        java.util.Arrays.asList(
                            new MutationStep(
                                "ok",
                                new WorkbookSelector.Current(),
                                new MutationAction.ClearWorkbookProtection()),
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
