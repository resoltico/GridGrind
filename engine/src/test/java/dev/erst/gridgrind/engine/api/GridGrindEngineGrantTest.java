package dev.erst.gridgrind.engine.api;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Locks the public Java engine boundary against implicit execution authority. */
class GridGrindEngineGrantTest {
  @TempDir Path root;

  @Test
  void rejectsMissingHostGrantAtThePublicInputBoundary() {
    assertThrows(
        NullPointerException.class,
        () -> new GridGrindRequestInputs(root, root.resolve("scratch"), null));
  }

  @Test
  void executesWhenThePublicInputCarriesAMatchingHostGrant() {
    ExecutionGrant grant =
        new ExecutionGrant(
            List.of(),
            List.of("ENSURE_SHEET"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum());

    WorkbookResult.Success success =
        assertInstanceOf(
            WorkbookResult.Success.class,
            GridGrindEngine.requestExecutor()
                .execute(
                    request(), new GridGrindRequestInputs(root, root.resolve("scratch"), grant)));

    assertEquals(
        1,
        assertInstanceOf(
                dev.erst.gridgrind.core.result.ExecutionJournal.Outcome.Succeeded.class,
                success.journal().outcome())
            .completedStepCount());
  }

  @Test
  void hostGrantRejectsMalformedAuthorityAndDeduplicatesExplicitValues() {
    ExecutionGrant grant =
        new ExecutionGrant(
            List.of(
                new ExecutionGrant.ReadAuthority.File(root.resolve("input.xlsx")),
                new ExecutionGrant.ReadAuthority.File(root.resolve("input.xlsx"))),
            List.of("ENSURE_SHEET", "ENSURE_SHEET"),
            new ExecutionGrant.WorkbookTargetAuthority.SelectedSheets(List.of("Budget", "Budget")),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum());

    assertEquals(1, grant.readableResources().size());
    assertEquals(List.of("ENSURE_SHEET"), grant.operationIds());
    assertEquals(
        List.of("Budget"),
        ((ExecutionGrant.WorkbookTargetAuthority.SelectedSheets) grant.targetAuthority())
            .sheetNames());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExecutionGrant(
                List.of(),
                List.of("ensure-sheet"),
                new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                AcceptancePolicy.minimum()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExecutionGrant.WorkbookTargetAuthority.SelectedSheets(List.of(" ")));
  }

  private static WorkbookPlan request() {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        List.of(
            new MutationStep(
                new StepId("ensure-sheet"),
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.EnsureSheet())));
  }
}
