package dev.erst.gridgrind.cli.examples;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.TaskStarterContract;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage for starter-contract validation and duplicate-id guards. */
class GridGrindTaskStarterPlansTest {
  @Test
  void taskStarterContractRejectsIncompatibleWorkspaceModes() {
    IllegalArgumentException selfContainedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "tasks/example.json",
                    ExampleWorkspaceMode.SELF_CONTAINED,
                    List.of("task-starter-assets/source.xlsx")));
    assertEquals(
        "SELF_CONTAINED task starters must not publish requiredPaths",
        selfContainedFailure.getMessage());

    IllegalArgumentException assetBackedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "tasks/example.json", ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS, List.of()));
    assertEquals(
        "REQUIRES_EXAMPLE_ASSETS task starters must publish requiredPaths",
        assetBackedFailure.getMessage());
  }

  @Test
  void taskStarterContractRequiresTaskScopedJsonPathsAndSupportsFactoryHelpers() {
    IllegalArgumentException wrongPrefix =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "examples/example.json", ExampleWorkspaceMode.SELF_CONTAINED, List.of()));
    assertEquals("suggestedRequestPath must start with tasks/", wrongPrefix.getMessage());

    IllegalArgumentException wrongSuffix =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "tasks/example.txt", ExampleWorkspaceMode.SELF_CONTAINED, List.of()));
    assertEquals("suggestedRequestPath must end with .json", wrongSuffix.getMessage());

    TaskStarterContract selfContained = TaskStarterContract.selfContained("tasks/example.json");
    assertEquals(ExampleWorkspaceMode.SELF_CONTAINED, selfContained.workspaceMode());
    assertEquals(List.of(), selfContained.requiredPaths());

    TaskStarterContract assetBacked =
        TaskStarterContract.assetBacked(
            "tasks/example.json", "task-starter-assets/source.xlsx", "payloads/data.xml");
    assertEquals(ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS, assetBacked.workspaceMode());
    assertEquals(
        List.of("task-starter-assets/source.xlsx", "payloads/data.xml"),
        assetBacked.requiredPaths());
  }

  @Test
  void taskStarterPlanRejectsBlankIds() {
    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of());

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterPlan(
                    " ", TaskStarterContract.selfContained("tasks/example.json"), plan));

    assertEquals("taskId must not be blank", failure.getMessage());
  }

  @Test
  void starterLookupRejectsUnknownIdsAndDuplicateMaps() {
    IllegalStateException missingContract =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindTaskStarterPlans.contractFor("NO_SUCH_TASK"));
    assertEquals("Missing task starter contract for NO_SUCH_TASK", missingContract.getMessage());

    IllegalStateException missingPlan =
        assertThrows(
            IllegalStateException.class, () -> GridGrindTaskStarterPlans.planFor("NO_SUCH_TASK"));
    assertEquals("Missing task starter plan for NO_SUCH_TASK", missingPlan.getMessage());
    assertEquals(
        "NEW",
        dev.erst.gridgrind.contract.json.GridGrindJson.requestTree(
                GridGrindTaskStarterPlans.planFor("DASHBOARD"))
            .path("source")
            .path("type")
            .stringValue());

    WorkbookPlan plan =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of());
    TaskStarterPlan left =
        new TaskStarterPlan(
            "DUPLICATE", TaskStarterContract.selfContained("tasks/left.json"), plan);
    TaskStarterPlan right =
        new TaskStarterPlan(
            "DUPLICATE", TaskStarterContract.selfContained("tasks/right.json"), plan);

    IllegalStateException duplicate =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindTaskStarterPlans.toStarterMap(List.of(left, right)));
    assertTrue(duplicate.getMessage().contains("Duplicate task starter plan for DUPLICATE"));
  }
}
