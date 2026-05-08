package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskKeywordMatchReport;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskPlanTemplate;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for English keyword matching over the CLI-owned task planner. */
class GridGrindTaskKeywordMatcherTest {
  @Test
  void keywordMatcherRanksDashboardAuditAndPivotTasksDeterministically() {
    TaskKeywordMatchReport dashboardReport =
        GridGrindTaskKeywordMatcher.reportFor("Create a monthly sales dashboard with charts");
    TaskKeywordMatchReport auditReport =
        GridGrindTaskKeywordMatcher.reportFor("Audit an existing workbook for health findings");
    TaskKeywordMatchReport pivotReport =
        GridGrindTaskKeywordMatcher.reportFor("build a pivot report from range data");

    assertEquals("DASHBOARD", dashboardReport.candidates().getFirst().task().id());
    assertTrue(dashboardReport.candidates().getFirst().matchedTerms().contains("dashboard"));
    assertTrue(dashboardReport.candidates().getFirst().matchedTerms().contains("chart"));
    assertTrue(dashboardReport.suggestedIntentTags().contains("dashboard"));
    assertFalse(dashboardReport.candidates().isEmpty());

    assertEquals("AUDIT_EXISTING_WORKBOOK", auditReport.candidates().getFirst().task().id());
    assertTrue(auditReport.candidates().getFirst().matchedTerms().contains("audit"));
    assertTrue(auditReport.candidates().getFirst().matchedTerms().contains("existing"));

    assertEquals("PIVOT_REPORT", pivotReport.candidates().getFirst().task().id());
    assertTrue(pivotReport.candidates().getFirst().matchedTerms().contains("pivot"));
    assertTrue(
        pivotReport.candidates().getFirst().starterTemplate().requestTemplate().steps().isEmpty());
  }

  @Test
  void keywordMatcherFindsCustomXmlAndWorkbookMaintenanceWorkflows() {
    TaskKeywordMatchReport customXmlReport =
        GridGrindTaskKeywordMatcher.reportFor("import custom xml mapping into an existing xlsx");
    TaskKeywordMatchReport maintenanceReport =
        GridGrindTaskKeywordMatcher.reportFor(
            "repair broken workbook comments and copy sheets safely");

    assertEquals("CUSTOM_XML_WORKFLOW", customXmlReport.candidates().getFirst().task().id());
    assertTrue(customXmlReport.candidates().getFirst().matchedTerms().contains("xml"));
    assertTrue(customXmlReport.candidates().getFirst().matchedTerms().contains("mapping"));
    assertTrue(
        customXmlReport.candidates().getFirst().reasons().stream()
            .anyMatch(reason -> reason.contains("capability summary")));

    assertEquals("WORKBOOK_MAINTENANCE", maintenanceReport.candidates().getFirst().task().id());
    assertTrue(maintenanceReport.candidates().getFirst().matchedTerms().contains("comment"));
    assertTrue(maintenanceReport.candidates().getFirst().matchedTerms().contains("copy"));
    assertFalse(maintenanceReport.candidates().isEmpty());
  }

  @Test
  void keywordMatcherLeavesUnmatchedVocabularyVisibleAndNormalizesPluralVariants() {
    TaskKeywordMatchReport noMatchReport =
        GridGrindTaskKeywordMatcher.reportFor("speaker notes presentation");
    TaskKeywordMatchReport normalizationReport =
        GridGrindTaskKeywordMatcher.reportFor(
            "___ boxes sizes quizzes matches classes brushes office");

    assertTrue(noMatchReport.candidates().isEmpty());
    assertTrue(noMatchReport.unmatchedTerms().contains("speaker"));
    assertTrue(noMatchReport.suggestedIntentTags().isEmpty());
    assertEquals(
        List.of("box", "size", "quiz", "match", "class", "brush"),
        normalizationReport.normalizedTerms());
    assertTrue(normalizationReport.unmatchedTerms().contains("quiz"));
    assertTrue(normalizationReport.unmatchedTerms().contains("class"));
    assertTrue(normalizationReport.unmatchedTerms().contains("brush"));
    assertTrue(normalizationReport.candidates().isEmpty());
  }

  @Test
  void keywordMatcherRejectsNullAndBlankQueries() {
    NullPointerException nullQuery =
        assertThrows(NullPointerException.class, () -> GridGrindTaskKeywordMatcher.reportFor(null));
    assertEquals("query must not be null", nullQuery.getMessage());

    IllegalArgumentException blankQuery =
        assertThrows(
            IllegalArgumentException.class, () -> GridGrindTaskKeywordMatcher.reportFor(" "));
    assertEquals("query must not be blank", blankQuery.getMessage());
  }

  @Test
  void candidateOrderingBreaksTiesByMatchedTermsThenTaskId() {
    TaskKeywordMatchReport.Candidate broaderMatch =
        candidate("DASHBOARD", 50, List.of("chart", "dashboard"));
    TaskKeywordMatchReport.Candidate narrowerMatch =
        candidate("PIVOT_REPORT", 50, List.of("chart"));
    List<TaskKeywordMatchReport.Candidate> byMatchedTerms =
        new ArrayList<>(List.of(narrowerMatch, broaderMatch));

    byMatchedTerms.sort(GridGrindTaskKeywordMatcher.candidateOrdering());

    assertEquals(List.of(broaderMatch, narrowerMatch), byMatchedTerms);

    TaskKeywordMatchReport.Candidate alphabeticallyLater =
        candidate("WORKBOOK_MAINTENANCE", 50, List.of("repair"));
    TaskKeywordMatchReport.Candidate alphabeticallyEarlier =
        candidate("CUSTOM_XML_WORKFLOW", 50, List.of("repair"));
    List<TaskKeywordMatchReport.Candidate> byTaskId =
        new ArrayList<>(List.of(alphabeticallyLater, alphabeticallyEarlier));

    byTaskId.sort(GridGrindTaskKeywordMatcher.candidateOrdering());

    assertEquals(List.of(alphabeticallyEarlier, alphabeticallyLater), byTaskId);
  }

  @Test
  void suggestedIntentTagsKeepsHighestScoreWhenDuplicateTagsReappear() {
    TaskEntry dashboardTask = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskPlanTemplate dashboardTemplate = GridGrindTaskPlanner.templateFor("DASHBOARD");
    List<String> suggestedTags =
        GridGrindTaskKeywordMatcher.suggestedIntentTags(
            List.of("dashboard"),
            List.of(
                new TaskKeywordMatchReport.Candidate(
                    dashboardTask, 40, List.of("dashboard"), List.of("high"), dashboardTemplate),
                new TaskKeywordMatchReport.Candidate(
                    dashboardTask, 20, List.of("dashboard"), List.of("low"), dashboardTemplate)));

    assertEquals(List.copyOf(new java.util.LinkedHashSet<>(suggestedTags)), suggestedTags);
    assertTrue(suggestedTags.contains("dashboard"));
    assertFalse(suggestedTags.isEmpty());
  }

  @Test
  void suggestedIntentTagsReplacesLowerScoreWhenStrongerDuplicateAppearsLater() {
    TaskEntry dashboardTask = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskPlanTemplate dashboardTemplate = GridGrindTaskPlanner.templateFor("DASHBOARD");
    List<String> suggestedTags =
        GridGrindTaskKeywordMatcher.suggestedIntentTags(
            List.of("dashboard"),
            List.of(
                new TaskKeywordMatchReport.Candidate(
                    dashboardTask, 20, List.of("dashboard"), List.of("low"), dashboardTemplate),
                new TaskKeywordMatchReport.Candidate(
                    dashboardTask, 40, List.of("dashboard"), List.of("high"), dashboardTemplate)));

    assertEquals(List.copyOf(new java.util.LinkedHashSet<>(suggestedTags)), suggestedTags);
    assertTrue(suggestedTags.contains("dashboard"));
    assertFalse(suggestedTags.isEmpty());
  }

  @Test
  void keywordMatcherProfileAndPhaseSurfaceTextCoversHiddenEnumBranches() {
    assertEquals(
        "overwrite in place",
        GridGrindTaskKeywordMatcher.persistenceModeSurface(TaskPersistenceMode.OVERWRITE_SOURCE));
    assertEquals(
        "export extract", GridGrindTaskKeywordMatcher.phasePurposeSurface(TaskPhasePurpose.EXPORT));
    assertEquals(
        "binary payload image object file",
        GridGrindTaskKeywordMatcher.inputKindSurface(TaskInputKind.BINARY_PAYLOAD));
    assertEquals(
        "drawing anchor position placement",
        GridGrindTaskKeywordMatcher.inputKindSurface(TaskInputKind.DRAWING_ANCHORS));
  }

  private static TaskKeywordMatchReport.Candidate candidate(
      String taskId, int score, List<String> matchedTerms) {
    return new TaskKeywordMatchReport.Candidate(
        GridGrindTaskCatalog.entryFor(taskId).orElseThrow(),
        score,
        matchedTerms,
        List.of("test"),
        GridGrindTaskPlanner.templateFor(taskId));
  }
}
