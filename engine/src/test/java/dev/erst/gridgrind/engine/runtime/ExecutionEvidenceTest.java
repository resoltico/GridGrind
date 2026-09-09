package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
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
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "ensure-budget",
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.EnsureSheet()),
                new AssertionStep(
                    "assert-budget",
                    new SheetSelector.ByName("Budget"),
                    new PresenceAssertion.SheetPresent()),
                new InspectionStep(
                    "inspect-budget",
                    new SheetSelector.ByName("Budget"),
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
