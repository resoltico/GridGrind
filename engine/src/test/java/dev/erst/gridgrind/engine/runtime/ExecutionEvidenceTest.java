package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies execution evidence projects each authored step kind and structural fact precisely. */
class ExecutionEvidenceTest {
  @Test
  void projectsMutationAssertionAndInspectionSemanticsIntoVerifiedExecutionEvidence() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("ensure-budget"),
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.EnsureSheet()),
                new AssertionStep(
                    new StepId("assert-budget"),
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new PresenceAssertion.SheetPresent()),
                new InspectionStep(
                    new StepId("inspect-budget"),
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new SheetIntrospectionQuery.GetSheetSummary())));

    WorkbookExecutionEvidence evidence =
        ExecutionEvidence.forExecution(
            request,
            new ExecutionInputBindings(
                Path.of("."),
                Path.of("tmp", "evidence"),
                ExecutionGrantTestSupport.noPublication()),
            CalculationReport.notRequested(),
            List.of(),
            List.of(),
            List.of(),
            true);

    assertEquals(
        List.of("ENSURE_SHEET", "EXPECT_SHEET_PRESENT", "GET_SHEET_SUMMARY"),
        evidence.admission().operations().stream()
            .map(WorkbookExecutionEvidence.Operation::operationId)
            .toList());
    assertInstanceOf(WorkbookExecutionEvidence.Structural.Established.class, evidence.structural());
    assertInstanceOf(
        WorkbookExecutionEvidence.Computational.NotRequested.class, evidence.computational());
    assertInstanceOf(
        WorkbookExecutionEvidence.Structural.NotEstablished.class,
        ExecutionEvidence.forExecution(
                request,
                new ExecutionInputBindings(
                    Path.of("."),
                    Path.of("tmp", "evidence-bound-unverified"),
                    ExecutionGrantTestSupport.noPublication()),
                CalculationReport.notRequested(),
                List.of(),
                List.of(),
                List.of(),
                false)
            .structural());
    assertInstanceOf(
        WorkbookExecutionEvidence.Structural.Established.class,
        ExecutionEvidence.forExecution(
                request, CalculationReport.notRequested(), List.of(), List.of(), List.of(), true)
            .structural());
    assertInstanceOf(
        WorkbookExecutionEvidence.Structural.NotEstablished.class,
        ExecutionEvidence.forExecution(
                request, CalculationReport.notRequested(), List.of(), List.of(), List.of(), false)
            .structural());
  }
}
