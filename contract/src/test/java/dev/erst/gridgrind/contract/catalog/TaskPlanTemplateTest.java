package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for task-plan templates emitted by the planner surface. */
class TaskPlanTemplateTest {
  @Test
  void constructorDefaultsProtocolVersionAndCopiesAuthoringNotes() {
    TaskEntry task =
        new TaskEntry(
            "TASK",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(new TaskCapabilityRef("sourceTypes", "NEW")),
                    List.of("note"))),
            List.of("pitfall"));
    WorkbookPlan requestTemplate =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of());
    List<String> mutableNotes = new ArrayList<>(List.of("note one"));

    TaskPlanTemplate defaultVersionTemplate =
        new TaskPlanTemplate(null, task, requestTemplate, mutableNotes);
    TaskPlanTemplate explicitVersionTemplate =
        new TaskPlanTemplate(
            GridGrindProtocolVersion.current(), task, requestTemplate, mutableNotes);
    mutableNotes.add("note two");

    assertEquals(GridGrindProtocolVersion.current(), defaultVersionTemplate.protocolVersion());
    assertEquals(List.of("note one"), defaultVersionTemplate.authoringNotes());
    assertEquals(GridGrindProtocolVersion.current(), explicitVersionTemplate.protocolVersion());
    assertThrows(
        UnsupportedOperationException.class,
        () -> defaultVersionTemplate.authoringNotes().add("nope"));
  }

  @Test
  void constructorRejectsNullTaskAndRequestTemplate() {
    WorkbookPlan requestTemplate =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of());
    TaskEntry task =
        new TaskEntry(
            "TASK",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(new TaskCapabilityRef("sourceTypes", "NEW")),
                    List.of("note"))),
            List.of("pitfall"));

    NullPointerException nullTask =
        assertThrows(
            NullPointerException.class,
            () -> new TaskPlanTemplate(null, null, requestTemplate, List.of("note")));
    assertEquals("task must not be null", nullTask.getMessage());

    NullPointerException nullRequestTemplate =
        assertThrows(
            NullPointerException.class,
            () -> new TaskPlanTemplate(null, task, null, List.of("note")));
    assertEquals("requestTemplate must not be null", nullRequestTemplate.getMessage());

    IllegalArgumentException blankNote =
        assertThrows(
            IllegalArgumentException.class,
            () -> new TaskPlanTemplate(null, task, requestTemplate, List.of(" ")));
    assertTrue(blankNote.getMessage().contains("authoringNotes"));
  }
}
