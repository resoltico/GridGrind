package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for starter task-plan scaffolds derived from task descriptors. */
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
    assertFalse(template.requestTemplate().steps().isEmpty());
    assertTrue(
        template.requestTemplate().steps().stream()
            .anyMatch(step -> "set-chart".equals(step.stepId())));
    assertTrue(
        template.authoringNotes().stream()
            .anyMatch(note -> note.contains("--print-protocol-catalog --search chart")));
  }

  @Test
  void plannerBuildsAuditStarterTemplate() {
    TaskPlanTemplate template = GridGrindTaskPlanner.templateFor("AUDIT_EXISTING_WORKBOOK");

    WorkbookPlan.WorkbookSource.ExistingFile source =
        assertInstanceOf(
            WorkbookPlan.WorkbookSource.ExistingFile.class, template.requestTemplate().source());
    assertTrue(source.path().endsWith(".xlsx"));
    assertTrue(source.path().contains("audit-input"));
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.None.class, template.requestTemplate().persistence());
    assertFalse(template.requestTemplate().steps().isEmpty());
    assertTrue(
        template.authoringNotes().stream().anyMatch(note -> note.contains("non-destructive")));
  }

  @Test
  void plannerBuildsStarterTemplatesForSpecializedExistingWorkbookTasks() {
    TaskPlanTemplate customXml = GridGrindTaskPlanner.templateFor("CUSTOM_XML_WORKFLOW");
    TaskPlanTemplate maintenance = GridGrindTaskPlanner.templateFor("WORKBOOK_MAINTENANCE");

    assertInstanceOf(
        WorkbookPlan.WorkbookSource.ExistingFile.class, customXml.requestTemplate().source());
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.SaveAs.class, customXml.requestTemplate().persistence());
    assertFalse(customXml.requestTemplate().steps().isEmpty());
    assertTrue(
        customXml.authoringNotes().stream().anyMatch(note -> note.contains("TODO_MAPPING_NAME")));

    assertInstanceOf(
        WorkbookPlan.WorkbookSource.ExistingFile.class, maintenance.requestTemplate().source());
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.SaveAs.class, maintenance.requestTemplate().persistence());
    assertFalse(maintenance.requestTemplate().steps().isEmpty());
    assertTrue(maintenance.authoringNotes().stream().anyMatch(note -> note.contains("Template")));
  }

  @Test
  void plannerRejectsUnknownAmbiguousAndIncompatibleTaskDefaults() {
    IllegalArgumentException unknownTask =
        assertThrows(
            IllegalArgumentException.class, () -> GridGrindTaskPlanner.templateFor("BOGUS_TASK"));
    assertTrue(unknownTask.getMessage().contains("Unknown task id"));

    TaskEntry ambiguousSourceTask =
        new TaskEntry(
            "AMBIGUOUS_SOURCE",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(
                        new TaskCapabilityRef("sourceTypes", "NEW"),
                        new TaskCapabilityRef("sourceTypes", "EXISTING"),
                        new TaskCapabilityRef("persistenceTypes", "SAVE_AS")),
                    List.of("note"))),
            List.of("pitfall"));
    IllegalStateException ambiguousSource =
        assertThrows(
            IllegalStateException.class, () -> GridGrindTaskPlanner.planFor(ambiguousSourceTask));
    assertTrue(ambiguousSource.getMessage().contains("multiple sourceTypes capabilities"));

    TaskEntry invalidOverwriteTask =
        new TaskEntry(
            "INVALID_OVERWRITE",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(
                        new TaskCapabilityRef("sourceTypes", "NEW"),
                        new TaskCapabilityRef("persistenceTypes", "OVERWRITE")),
                    List.of("note"))),
            List.of("pitfall"));
    IllegalStateException invalidOverwrite =
        assertThrows(
            IllegalStateException.class, () -> GridGrindTaskPlanner.planFor(invalidOverwriteTask));
    assertTrue(invalidOverwrite.getMessage().contains("cannot plan OVERWRITE persistence"));
  }

  @Test
  void plannerBuildsOverwriteStarterTemplateForExistingWorkbookTasks() {
    TaskEntry overwriteTask =
        new TaskEntry(
            "OVERWRITE_EXISTING",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(
                        new TaskCapabilityRef("sourceTypes", "EXISTING"),
                        new TaskCapabilityRef("persistenceTypes", "OVERWRITE")),
                    List.of("note"))),
            List.of("pitfall"));

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
        new TaskEntry(
            "AD_HOC_DISCOVERY",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(
                        new TaskCapabilityRef("sourceTypes", "NEW"),
                        new TaskCapabilityRef("persistenceTypes", "NONE")),
                    List.of("note"))),
            List.of("pitfall"));
    TaskEntry saveAsTask =
        new TaskEntry(
            "AD_HOC_EXPORT",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(
                        new TaskCapabilityRef("sourceTypes", "EXISTING"),
                        new TaskCapabilityRef("persistenceTypes", "SAVE_AS")),
                    List.of("note"))),
            List.of("pitfall"));

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
  void plannerRejectsNullAndUnsupportedCapabilityDefaults() {
    NullPointerException nullTask =
        assertThrows(NullPointerException.class, () -> GridGrindTaskPlanner.planFor(null));
    assertEquals("task must not be null", nullTask.getMessage());

    TaskEntry unsupportedSourceTask =
        new TaskEntry(
            "UNSUPPORTED_SOURCE",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(
                        new TaskCapabilityRef("sourceTypes", "ARCHIVE"),
                        new TaskCapabilityRef("persistenceTypes", "NONE")),
                    List.of("note"))),
            List.of("pitfall"));
    IllegalStateException unsupportedSource =
        assertThrows(
            IllegalStateException.class, () -> GridGrindTaskPlanner.planFor(unsupportedSourceTask));
    assertTrue(unsupportedSource.getMessage().contains("unsupported source type"));

    TaskEntry unsupportedPersistenceTask =
        new TaskEntry(
            "UNSUPPORTED_PERSISTENCE",
            "summary",
            List.of("office"),
            List.of("outcome"),
            List.of("input"),
            List.of("feature"),
            List.of(
                new TaskPhase(
                    "Phase",
                    "Objective",
                    List.of(
                        new TaskCapabilityRef("sourceTypes", "NEW"),
                        new TaskCapabilityRef("persistenceTypes", "ARCHIVE")),
                    List.of("note"))),
            List.of("pitfall"));
    IllegalStateException unsupportedPersistence =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindTaskPlanner.planFor(unsupportedPersistenceTask));
    assertTrue(unsupportedPersistence.getMessage().contains("unsupported persistence type"));

    TaskEntry missingPersistenceTask =
        new TaskEntry(
            "MISSING_PERSISTENCE",
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
    IllegalStateException missingPersistence =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindTaskPlanner.planFor(missingPersistenceTask));
    assertTrue(missingPersistence.getMessage().contains("does not declare any persistenceTypes"));
  }
}
