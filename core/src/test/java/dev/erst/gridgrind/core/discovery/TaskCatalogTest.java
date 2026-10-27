package dev.erst.gridgrind.core.discovery;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import org.junit.jupiter.api.Test;

/** Validates product task definitions and exact capability references. */
class TaskCatalogTest {
  @Test
  void taskCatalogValidationRejectsDuplicateIdsAndDanglingReferences() {
    TaskPhase newSourcePhase =
        TaskTestFixtures.phase(
            java.util.List.of(new TaskCapabilityRef("workbookSourceTypes", "NEW")));
    TaskEntry left =
        new TaskEntry(
            "DUPLICATE",
            java.util.List.of("office"),
            TaskTestFixtures.discoveryProfile("duplicate"),
            TaskTestFixtures.narrative("one"),
            profile(),
            TaskTestFixtures.interactionProfile(),
            TaskStarterContract.selfContained("duplicate-request.json"),
            TaskTestFixtures.workflow(java.util.List.of(newSourcePhase)));
    TaskEntry right =
        new TaskEntry(
            "DUPLICATE",
            java.util.List.of("office"),
            TaskTestFixtures.discoveryProfile("duplicate"),
            TaskTestFixtures.narrative("two"),
            profile(),
            TaskTestFixtures.interactionProfile(),
            TaskStarterContract.selfContained("duplicate-request.json"),
            TaskTestFixtures.workflow(
                java.util.List.of(
                    new TaskPhase(
                        TaskPhasePurpose.AUTHOR,
                        "Phase Two",
                        "Objective",
                        java.util.List.of(new TaskCapabilityRef("workbookSourceTypes", "NEW")),
                        java.util.List.of("note")))));
    IllegalArgumentException duplicateTasks =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskCatalog(
                    dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT,
                    java.util.List.of(left, right)));
    assertEquals("tasks must not contain duplicate DUPLICATE", duplicateTasks.getMessage());

    IllegalArgumentException duplicateCapabilities =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskPhase(
                    TaskPhasePurpose.AUTHOR,
                    "Phase",
                    "Objective",
                    java.util.List.of(
                        new TaskCapabilityRef("workbookSourceTypes", "NEW"),
                        new TaskCapabilityRef("workbookSourceTypes", "NEW")),
                    java.util.List.of("note")));
    assertEquals(
        "capabilityRefs must not contain duplicate workbookSourceTypes:NEW",
        duplicateCapabilities.getMessage());

    IllegalArgumentException emptyCapabilities =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskPhase(
                    TaskPhasePurpose.AUTHOR,
                    "Phase",
                    "Objective",
                    java.util.List.of(),
                    java.util.List.of("note")));
    assertEquals("capabilityRefs must not be empty", emptyCapabilities.getMessage());

    IllegalArgumentException emptyPhases =
        assertThrows(
            IllegalArgumentException.class,
            () -> new TaskWorkflow(java.util.List.of(), java.util.List.of("pitfall")));
    assertEquals("phases must not be empty", emptyPhases.getMessage());

    IllegalStateException danglingReference =
        assertThrows(
            IllegalStateException.class,
            () ->
                TaskCatalogs.validateCapabilityReferences(
                    new TaskCatalog(
                        dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT,
                        java.util.List.of(brokenTask()))));
    assertTrue(
        danglingReference
            .getMessage()
            .contains("Task BROKEN references unknown protocol capability"));
    assertEquals(
        "protocolVersion must not be null",
        assertThrows(
                NullPointerException.class, () -> new TaskCatalog(null, java.util.List.of(left)))
            .getMessage());
  }

  @Test
  void taskCatalogValidationAcceptsTheBuiltInCatalog() {
    assertDoesNotThrow(() -> TaskCatalogs.validateCapabilityReferences(TaskCatalogs.catalog()));
  }

  @Test
  void examplePlanSupportBuildsSheetScopedTableSelectors() {
    var table =
        new dev.erst.gridgrind.core.plan.TableSelector.ByNameOnSheet(
            new TableName("BudgetTable"), new SheetName("Budget"));

    assertEquals("BudgetTable", table.name().value());
    assertEquals("Budget", table.sheetName().value());
  }

  private static TaskExecutionProfile profile() {
    return TaskTestFixtures.profile();
  }

  private static TaskEntry brokenTask() {
    return TaskTestFixtures.task(
        "BROKEN",
        profile(),
        java.util.List.of(
            TaskTestFixtures.phase(
                java.util.List.of(
                    new TaskCapabilityRef("mutationActionTypes", "NO_SUCH_ACTION")))));
  }
}
