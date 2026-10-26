package dev.erst.gridgrind.engine.api;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

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
    GridGrindExecutionGrant grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of("ENSURE_SHEET"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            GridGrindHostAcceptancePolicy.minimum());

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
    GridGrindExecutionGrant.Bounded grant =
        new GridGrindExecutionGrant.Bounded(
            List.of(
                new GridGrindExecutionGrant.ReadAuthority.File(root.resolve("input.xlsx")),
                new GridGrindExecutionGrant.ReadAuthority.File(root.resolve("input.xlsx"))),
            List.of("ENSURE_SHEET", "ENSURE_SHEET"),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets(
                List.of("Budget", "Budget")),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            GridGrindHostAcceptancePolicy.minimum());

    assertEquals(1, grant.readableResources().size());
    assertEquals(List.of("ENSURE_SHEET"), grant.operationIds());
    assertEquals(
        List.of("Budget"),
        ((GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets)
                grant.workbookTargetAuthority())
            .sheetNames());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new GridGrindExecutionGrant.Bounded(
                List.of(),
                List.of("ensure-sheet"),
                new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new GridGrindExecutionGrant.PublicationAuthority.None(),
                List.of(),
                GridGrindHostAcceptancePolicy.minimum()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets(List.of(" ")));
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
