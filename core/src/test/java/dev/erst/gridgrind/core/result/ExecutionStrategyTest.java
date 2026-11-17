package dev.erst.gridgrind.core.result;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Checks concrete strategy evidence and full-workbook materialization requirements. */
class ExecutionStrategyTest {
  @Test
  void reportsFullViewsRequiredByHostAcceptanceAndStagedVerification() {
    var event =
        request(
            new ExecutionModeSpec.EventRead(),
            new WorkbookPlan.WorkbookSource.ExistingFile("source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None());
    assertEquals(
        WorkbookExecutionEvidence.Backend.EVENT_READ,
        BackendCapabilities.executionStrategy(event, ExecutionGrant.denied()).backend());
    assertFalse(
        BackendCapabilities.executionStrategy(event, ExecutionGrant.denied())
            .fullWorkbookMaterializationRequired());
    var required =
        new ExecutionGrant(
            0,
            List.of(),
            List.of(),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            new AcceptancePolicy.RequireAll(
                List.of(
                    new AcceptancePolicy.Requirement.TerminalAssertions(
                        List.of(
                            new AssertionStep(
                                new StepId("host-sheet"),
                                new SheetSelector.ByName(new SheetName("Data")),
                                new PresenceAssertion.SheetPresent()))))));
    assertTrue(
        BackendCapabilities.executionStrategy(event, required)
            .fullWorkbookMaterializationRequired());
    var streaming =
        request(
            new ExecutionModeSpec.StreamingWrite(),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None());
    assertFalse(
        BackendCapabilities.executionStrategy(streaming, ExecutionGrant.denied())
            .fullWorkbookMaterializationRequired());
    var saved =
        request(
            new ExecutionModeSpec.StreamingWrite(),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "saved.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT));
    assertTrue(
        BackendCapabilities.executionStrategy(saved, ExecutionGrant.denied())
            .fullWorkbookMaterializationRequired());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookExecutionEvidence.ExecutionStrategy(
                WorkbookExecutionEvidence.Backend.FULL_XSSF, false));
  }

  private static WorkbookPlan request(
      ExecutionModeSpec mode,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.mode(mode),
        FormulaEnvironmentSpec.empty(),
        List.of());
  }
}
