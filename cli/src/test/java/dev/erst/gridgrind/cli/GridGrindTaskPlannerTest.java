package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskCapabilityRef;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskExecutionProfile;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhase;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskPlanTemplate;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for CLI-owned starter task-plan scaffolds derived from task descriptors. */
class GridGrindTaskPlannerTest {
  @Test
  void plannerBuildsDashboardStarterTemplate() {
    TaskPlanTemplate template = GridGrindTaskPlanner.templateFor("DASHBOARD");

    assertEquals("DASHBOARD", template.task().id());
    assertInstanceOf(WorkbookPlan.WorkbookSource.New.class, template.requestTemplate().source());
    WorkbookPlan.WorkbookPersistence.SaveAs persistence =
        assertInstanceOf(
            WorkbookPlan.WorkbookPersistence.SaveAs.class,
            template.requestTemplate().persistence());
    assertTrue(persistence.path().endsWith(".xlsx"));
    assertTrue(persistence.path().contains("dashboard"));
    assertTrue(template.requestTemplate().steps().isEmpty());
    assertTrue(
        template.authoringNotes().stream()
            .anyMatch(note -> note.contains("source and persistence are scaffolded")));
    assertTrue(
        template.authoringNotes().stream().anyMatch(note -> note.contains("output workbook path")));
  }

  @Test
  void plannerBuildsAuditStarterTemplate() {
    TaskPlanTemplate template = GridGrindTaskPlanner.templateFor("AUDIT_EXISTING_WORKBOOK");

    WorkbookPlan.WorkbookSource.ExistingFile source =
        assertInstanceOf(
            WorkbookPlan.WorkbookSource.ExistingFile.class, template.requestTemplate().source());
    assertTrue(source.path().endsWith(".xlsx"));
    assertTrue(source.path().contains("audit-existing-workbook"));
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.None.class, template.requestTemplate().persistence());
    assertTrue(template.requestTemplate().steps().isEmpty());
    assertTrue(
        template.authoringNotes().stream().anyMatch(note -> note.contains("non-destructive")));
  }

  @Test
  void plannerCarriesTaskPitfallsIntoSpecializedExistingWorkbookTemplates() {
    TaskPlanTemplate customXml = GridGrindTaskPlanner.templateFor("CUSTOM_XML_WORKFLOW");
    TaskPlanTemplate maintenance = GridGrindTaskPlanner.templateFor("WORKBOOK_MAINTENANCE");

    assertInstanceOf(
        WorkbookPlan.WorkbookSource.ExistingFile.class, customXml.requestTemplate().source());
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.SaveAs.class, customXml.requestTemplate().persistence());
    assertTrue(customXml.requestTemplate().steps().isEmpty());
    assertTrue(
        customXml.authoringNotes().stream()
            .anyMatch(note -> note.contains("IMPORT_CUSTOM_XML_MAPPING requires an existing")));

    assertInstanceOf(
        WorkbookPlan.WorkbookSource.ExistingFile.class, maintenance.requestTemplate().source());
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.SaveAs.class, maintenance.requestTemplate().persistence());
    assertTrue(maintenance.requestTemplate().steps().isEmpty());
    assertTrue(
        maintenance.authoringNotes().stream()
            .anyMatch(note -> note.contains("copy-sheet requires an existing source sheet")));
  }

  @Test
  void plannerRejectsUnknownAndIncompatibleTaskDefaults() {
    IllegalArgumentException unknownTask =
        assertThrows(
            IllegalArgumentException.class, () -> GridGrindTaskPlanner.templateFor("BOGUS_TASK"));
    assertTrue(unknownTask.getMessage().contains("Unknown task id"));

    TaskEntry invalidOverwriteTask =
        task(
            "INVALID_OVERWRITE",
            new TaskExecutionProfile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.OVERWRITE_SOURCE,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")))));
    IllegalStateException invalidOverwrite =
        assertThrows(
            IllegalStateException.class, () -> GridGrindTaskPlanner.planFor(invalidOverwriteTask));
    assertTrue(invalidOverwrite.getMessage().contains("cannot plan OVERWRITE persistence"));
  }

  @Test
  void plannerRejectsNullAndBlankTaskIds() {
    NullPointerException nullTaskId =
        assertThrows(NullPointerException.class, () -> GridGrindTaskPlanner.templateFor(null));
    assertEquals("taskId must not be null", nullTaskId.getMessage());

    IllegalArgumentException blankTaskId =
        assertThrows(IllegalArgumentException.class, () -> GridGrindTaskPlanner.templateFor(" "));
    assertEquals("taskId must not be blank", blankTaskId.getMessage());
  }

  @Test
  void plannerBuildsOverwriteStarterTemplateForExistingWorkbookTasks() {
    TaskEntry overwriteTask =
        task(
            "OVERWRITE_EXISTING",
            new TaskExecutionProfile(
                TaskSourceMode.EXISTING_WORKBOOK,
                TaskPersistenceMode.OVERWRITE_SOURCE,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")))));

    TaskPlanTemplate template = GridGrindTaskPlanner.planFor(overwriteTask);

    WorkbookPlan.WorkbookSource.ExistingFile source =
        assertInstanceOf(
            WorkbookPlan.WorkbookSource.ExistingFile.class, template.requestTemplate().source());
    assertTrue(source.path().endsWith(".xlsx"));
    assertTrue(source.path().contains("overwrite-existing"));
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.OverwriteSource.class,
        template.requestTemplate().persistence());
  }

  @Test
  void plannerBuildsGenericTemplatesForAdHocNoneAndSaveAsTasks() {
    TaskEntry noneTask =
        task(
            "AD_HOC_DISCOVERY",
            new TaskExecutionProfile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.NONE,
                TaskMutationMode.READ_ONLY,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("inspectionQueryTypes", "GET_CELLS")))));
    TaskEntry saveAsTask =
        task(
            "AD_HOC_EXPORT",
            new TaskExecutionProfile(
                TaskSourceMode.EXISTING_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")))));

    TaskPlanTemplate noneTemplate = GridGrindTaskPlanner.planFor(noneTask);
    TaskPlanTemplate saveAsTemplate = GridGrindTaskPlanner.planFor(saveAsTask);
    String noneNotes = String.join("\n", noneTemplate.authoringNotes());
    String saveAsNotes = String.join("\n", saveAsTemplate.authoringNotes());

    assertInstanceOf(
        WorkbookPlan.WorkbookSource.New.class, noneTemplate.requestTemplate().source());
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.None.class, noneTemplate.requestTemplate().persistence());
    assertTrue(noneTemplate.requestTemplate().steps().isEmpty());
    assertTrue(noneNotes.contains("non-destructive"));

    WorkbookPlan.WorkbookSource.ExistingFile existingSource =
        assertInstanceOf(
            WorkbookPlan.WorkbookSource.ExistingFile.class,
            saveAsTemplate.requestTemplate().source());
    WorkbookPlan.WorkbookPersistence.SaveAs saveAs =
        assertInstanceOf(
            WorkbookPlan.WorkbookPersistence.SaveAs.class,
            saveAsTemplate.requestTemplate().persistence());
    assertEquals("todo-ad-hoc-export-input.xlsx", existingSource.path());
    assertEquals("todo-ad-hoc-export-output.xlsx", saveAs.path());
    assertTrue(saveAsTemplate.requestTemplate().steps().isEmpty());
    assertFalse(saveAsNotes.contains("non-destructive"));
  }

  @Test
  void plannerRejectsNullTasksAndCarriesExternalAssetNotes() {
    NullPointerException nullTask =
        assertThrows(NullPointerException.class, () -> GridGrindTaskPlanner.planFor(null));
    assertEquals("task must not be null", nullTask.getMessage());

    TaskPlanTemplate externalPayloadTemplate =
        GridGrindTaskPlanner.planFor(
            task(
                "EXTERNAL_PAYLOAD_IMPORT",
                new TaskExecutionProfile(
                    TaskSourceMode.EXISTING_WORKBOOK,
                    TaskPersistenceMode.SAVE_AS,
                    TaskMutationMode.MUTATING,
                    TaskAssetMode.REQUIRES_EXTERNAL_PAYLOADS),
                List.of(
                    phase(
                        List.of(
                            new TaskCapabilityRef(
                                "mutationActionTypes", "IMPORT_CUSTOM_XML_MAPPING"))))));

    assertTrue(
        externalPayloadTemplate.authoringNotes().stream()
            .anyMatch(note -> note.contains("external payload files")));
  }

  private static TaskEntry task(
      String id, TaskExecutionProfile executionProfile, List<TaskPhase> phases) {
    return new TaskEntry(
        id,
        "summary",
        executionProfile,
        List.of(),
        List.of(),
        List.of("office"),
        List.of("outcome"),
        List.of("input"),
        List.of("feature"),
        phases,
        List.of("pitfall"));
  }

  private static TaskPhase phase(List<TaskCapabilityRef> capabilityRefs) {
    return new TaskPhase(
        TaskPhasePurpose.AUTHOR, "Phase", "Objective", capabilityRefs, List.of("note"));
  }
}
