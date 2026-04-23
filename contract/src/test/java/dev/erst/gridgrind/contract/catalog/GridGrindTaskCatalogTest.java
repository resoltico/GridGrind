package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import org.junit.jupiter.api.Test;

/** Tests for the high-level task catalog layered on top of the exact protocol catalog. */
class GridGrindTaskCatalogTest {
  @Test
  void exposesDeterministicTaskCatalogEntries() throws IOException {
    TaskCatalog catalog = GridGrindTaskCatalog.catalog();
    TaskCatalog decoded =
        GridGrindJson.readTaskCatalog(GridGrindJson.writeTaskCatalogBytes(catalog));

    assertFalse(catalog.tasks().isEmpty());
    assertEquals(catalog, decoded);
    assertTrue(GridGrindTaskCatalog.entryFor("DASHBOARD").isPresent());
    assertEquals(
        "inspectionQueryTypes:GET_CHARTS",
        GridGrindTaskCatalog.entryFor("DASHBOARD")
            .orElseThrow()
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
    TaskEntry left =
        new TaskEntry(
            "DUPLICATE",
            "one",
            java.util.List.of("office"),
            java.util.List.of("outcome"),
            java.util.List.of("input"),
            java.util.List.of("feature"),
            java.util.List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    java.util.List.of(new TaskCapabilityRef("sourceTypes", "NEW")),
                    java.util.List.of("note"))),
            java.util.List.of("pitfall"));
    TaskEntry right =
        new TaskEntry(
            "DUPLICATE",
            "two",
            java.util.List.of("office"),
            java.util.List.of("outcome"),
            java.util.List.of("input"),
            java.util.List.of("feature"),
            java.util.List.of(
                new TaskPhase(
                    "Phase Two",
                    "Objective",
                    java.util.List.of(new TaskCapabilityRef("sourceTypes", "NEW")),
                    java.util.List.of("note"))),
            java.util.List.of("pitfall"));
    IllegalStateException duplicateTasks =
        assertThrows(
            IllegalStateException.class,
            () -> new TaskCatalog(null, java.util.List.of(left, right)));
    assertTrue(duplicateTasks.getMessage().contains("Duplicate tasks detected"));

    IllegalStateException duplicateCapabilities =
        assertThrows(
            IllegalStateException.class,
            () ->
                new TaskPhase(
                    "Phase",
                    "Objective",
                    java.util.List.of(
                        new TaskCapabilityRef("sourceTypes", "NEW"),
                        new TaskCapabilityRef("sourceTypes", "NEW")),
                    java.util.List.of("note")));
    assertTrue(duplicateCapabilities.getMessage().contains("Duplicate capabilityRefs detected"));

    IllegalArgumentException emptyCapabilities =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskPhase(
                    "Phase", "Objective", java.util.List.of(), java.util.List.of("note")));
    assertEquals("capabilityRefs must not be empty", emptyCapabilities.getMessage());

    IllegalArgumentException emptyPhases =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskEntry(
                    "TASK",
                    "summary",
                    java.util.List.of("office"),
                    java.util.List.of("outcome"),
                    java.util.List.of("input"),
                    java.util.List.of("feature"),
                    java.util.List.of(),
                    java.util.List.of("pitfall")));
    assertEquals("phases must not be empty", emptyPhases.getMessage());

    IllegalStateException danglingReference =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindTaskCatalog.validateCapabilityReferences(
                    new TaskCatalog(
                        null,
                        java.util.List.of(
                            new TaskEntry(
                                "BROKEN",
                                "summary",
                                java.util.List.of("office"),
                                java.util.List.of("outcome"),
                                java.util.List.of("input"),
                                java.util.List.of("feature"),
                                java.util.List.of(
                                    new TaskPhase(
                                        "Phase",
                                        "Objective",
                                        java.util.List.of(
                                            new TaskCapabilityRef(
                                                "mutationActionTypes", "NO_SUCH_ACTION")),
                                        java.util.List.of("note"))),
                                java.util.List.of("pitfall"))))));
    assertTrue(
        danglingReference
            .getMessage()
            .contains("Task BROKEN references unknown protocol capability"));
  }

  @Test
  void taskCatalogValidationAcceptsTheBuiltInCatalog() {
    assertDoesNotThrow(
        () -> GridGrindTaskCatalog.validateCapabilityReferences(GridGrindTaskCatalog.catalog()));
  }
}
