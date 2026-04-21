package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
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
    assertTrue(template.requestTemplate().steps().isEmpty());
    assertTrue(
        template.authoringNotes().stream()
            .anyMatch(note -> note.contains("--print-protocol-catalog --operation <group>:<id>")));
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
    assertTrue(
        template.authoringNotes().stream().anyMatch(note -> note.contains("non-destructive")));
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

  @Test
  void goalPlannerRanksDashboardAndAuditTasksDeterministically() {
    GoalPlanReport dashboardReport =
        GridGrindGoalPlanner.reportFor("Create a monthly sales dashboard with charts");
    GoalPlanReport auditReport =
        GridGrindGoalPlanner.reportFor("Audit an existing workbook for health findings");
    GoalPlanReport noMatchReport = GridGrindGoalPlanner.reportFor("speaker notes presentation");

    assertEquals("DASHBOARD", dashboardReport.candidates().getFirst().task().id());
    assertTrue(dashboardReport.candidates().getFirst().matchedTerms().contains("dashboard"));
    assertTrue(dashboardReport.candidates().getFirst().matchedTerms().contains("chart"));
    assertEquals("AUDIT_EXISTING_WORKBOOK", auditReport.candidates().getFirst().task().id());
    assertTrue(auditReport.candidates().getFirst().matchedTerms().contains("audit"));
    assertTrue(auditReport.candidates().getFirst().matchedTerms().contains("existing"));
    assertTrue(noMatchReport.candidates().isEmpty());
    assertTrue(noMatchReport.unmatchedTerms().contains("speaker"));
    assertTrue(noMatchReport.suggestedIntentTags().contains("dashboard"));
  }

  @Test
  void goalPlannerNormalizesPluralVariantsAndBreaksTiesByStableTaskOrdering() {
    GoalPlanReport tieBreakReport = GridGrindGoalPlanner.reportFor("saved");
    GoalPlanReport normalizationReport =
        GridGrindGoalPlanner.reportFor("___ boxes sizes quizzes matches classes brushes office");

    assertEquals(
        List.of("DASHBOARD", "DATA_ENTRY_WORKFLOW", "TABULAR_REPORT"),
        tieBreakReport.candidates().stream().map(candidate -> candidate.task().id()).toList());
    assertEquals(
        List.of("box", "size", "quiz", "match", "class", "brush"),
        normalizationReport.normalizedTerms());
    assertTrue(
        normalizationReport.unmatchedTerms().containsAll(normalizationReport.normalizedTerms()));
  }
}
