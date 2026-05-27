package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import org.junit.jupiter.api.Test;

/** Tests for the high-level task catalog layered on top of the exact protocol catalog. */
class GridGrindTaskCatalogTest {
  @Test
  void exposesDeterministicTaskCatalogEntries() throws IOException {
    TaskCatalog catalog = GridGrindTaskCatalog.catalog();
    TaskCatalog decoded =
        GridGrindCliJson.readTaskCatalog(GridGrindCliJson.writeTaskCatalogBytes(catalog));

    assertFalse(catalog.tasks().isEmpty());
    assertEquals(catalog, decoded);
    assertTrue(GridGrindTaskCatalog.entryFor("DASHBOARD").isPresent());
    assertEquals(
        "inspectionQueryTypes:GET_CHARTS",
        GridGrindTaskCatalog.entryFor("DASHBOARD")
            .orElseThrow()
            .workflow()
            .phases()
            .getLast()
            .capabilityRefs()
            .getFirst()
            .qualifiedId());
    assertTrue(GridGrindTaskCatalog.entryFor("TABULAR_REPORT").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("AUDIT_EXISTING_WORKBOOK").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("DATA_ENTRY_WORKFLOW").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("CUSTOM_XML_WORKFLOW").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("PIVOT_REPORT").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("DRAWING_AND_SIGNATURE_WORKFLOW").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("WORKBOOK_MAINTENANCE").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("BOGUS_TASK").isEmpty());
  }

  @Test
  void taskCatalogValidationRejectsDuplicateIdsAndDanglingReferences() {
    TaskPhase newSourcePhase =
        TaskTestFixtures.phase(java.util.List.of(new TaskCapabilityRef("sourceTypes", "NEW")));
    TaskEntry left =
        new TaskEntry(
            "DUPLICATE",
            TaskTestFixtures.discoveryProfile("duplicate"),
            TaskTestFixtures.narrative("one"),
            profile(),
            TaskTestFixtures.interactionProfile(),
            TaskStarterContract.selfContained("tasks/duplicate-request.json"),
            TaskTestFixtures.workflow(java.util.List.of(newSourcePhase)));
    TaskEntry right =
        new TaskEntry(
            "DUPLICATE",
            TaskTestFixtures.discoveryProfile("duplicate"),
            TaskTestFixtures.narrative("two"),
            profile(),
            TaskTestFixtures.interactionProfile(),
            TaskStarterContract.selfContained("tasks/duplicate-request.json"),
            TaskTestFixtures.workflow(
                java.util.List.of(
                    new TaskPhase(
                        TaskPhasePurpose.AUTHOR,
                        "Phase Two",
                        "Objective",
                        java.util.List.of(new TaskCapabilityRef("sourceTypes", "NEW")),
                        java.util.List.of("note")))));
    IllegalArgumentException duplicateTasks =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskCatalog(
                    dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
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
                        new TaskCapabilityRef("sourceTypes", "NEW"),
                        new TaskCapabilityRef("sourceTypes", "NEW")),
                    java.util.List.of("note")));
    assertEquals(
        "capabilityRefs must not contain duplicate sourceTypes:NEW",
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
                GridGrindTaskCatalog.validateCapabilityReferences(
                    new TaskCatalog(
                        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
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
    assertDoesNotThrow(
        () -> GridGrindTaskCatalog.validateCapabilityReferences(GridGrindTaskCatalog.catalog()));
  }

  @Test
  void examplePlanSupportBuildsSheetScopedTableSelectors() {
    var table =
        new dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet(
            "BudgetTable", "Budget");

    assertEquals("BudgetTable", table.name());
    assertEquals("Budget", table.sheetName());
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
