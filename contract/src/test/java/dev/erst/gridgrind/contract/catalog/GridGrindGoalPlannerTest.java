package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for freeform-goal ranking over the contract-owned task catalog. */
class GridGrindGoalPlannerTest {
  @Test
  void goalPlannerRanksDashboardAuditAndPivotTasksDeterministically() {
    GoalPlanReport dashboardReport =
        GridGrindGoalPlanner.reportFor("Create a monthly sales dashboard with charts");
    GoalPlanReport auditReport =
        GridGrindGoalPlanner.reportFor("Audit an existing workbook for health findings");
    GoalPlanReport pivotReport =
        GridGrindGoalPlanner.reportFor("build a pivot report from range data");

    assertEquals("DASHBOARD", dashboardReport.candidates().getFirst().task().id());
    assertTrue(dashboardReport.candidates().getFirst().matchedTerms().contains("dashboard"));
    assertTrue(dashboardReport.candidates().getFirst().matchedTerms().contains("chart"));
    assertEquals(
        List.of("charts", "dashboard", "kpi", "office", "summary"),
        dashboardReport.suggestedIntentTags());
    assertEquals(1, dashboardReport.candidates().size());

    assertEquals("AUDIT_EXISTING_WORKBOOK", auditReport.candidates().getFirst().task().id());
    assertTrue(auditReport.candidates().getFirst().matchedTerms().contains("audit"));
    assertTrue(auditReport.candidates().getFirst().matchedTerms().contains("existing"));

    assertEquals("PIVOT_REPORT", pivotReport.candidates().getFirst().task().id());
    assertTrue(pivotReport.candidates().getFirst().matchedTerms().contains("pivot"));
    assertFalse(
        pivotReport.candidates().getFirst().starterTemplate().requestTemplate().steps().isEmpty());
  }

  @Test
  void goalPlannerNowFindsCustomXmlAndWorkbookMaintenanceWorkflows() {
    GoalPlanReport customXmlReport =
        GridGrindGoalPlanner.reportFor("import custom xml mapping into an existing xlsx");
    GoalPlanReport maintenanceReport =
        GridGrindGoalPlanner.reportFor("repair broken workbook comments and copy sheets safely");

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
  void goalPlannerLeavesUnmatchedVocabularyVisibleAndNormalizesPluralVariants() {
    GoalPlanReport noMatchReport = GridGrindGoalPlanner.reportFor("speaker notes presentation");
    GoalPlanReport normalizationReport =
        GridGrindGoalPlanner.reportFor("___ boxes sizes quizzes matches classes brushes office");

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
  void candidateOrderingBreaksTiesByMatchedTermsThenTaskId() {
    GoalPlanReport.Candidate broaderMatch =
        candidate("DASHBOARD", 50, List.of("chart", "dashboard"));
    GoalPlanReport.Candidate narrowerMatch = candidate("PIVOT_REPORT", 50, List.of("chart"));
    List<GoalPlanReport.Candidate> byMatchedTerms =
        new ArrayList<>(List.of(narrowerMatch, broaderMatch));

    byMatchedTerms.sort(GridGrindGoalPlanner.candidateOrdering());

    assertEquals(List.of(broaderMatch, narrowerMatch), byMatchedTerms);

    GoalPlanReport.Candidate alphabeticallyLater =
        candidate("WORKBOOK_MAINTENANCE", 50, List.of("repair"));
    GoalPlanReport.Candidate alphabeticallyEarlier =
        candidate("CUSTOM_XML_WORKFLOW", 50, List.of("repair"));
    List<GoalPlanReport.Candidate> byTaskId =
        new ArrayList<>(List.of(alphabeticallyLater, alphabeticallyEarlier));

    byTaskId.sort(GridGrindGoalPlanner.candidateOrdering());

    assertEquals(List.of(alphabeticallyEarlier, alphabeticallyLater), byTaskId);
  }

  @Test
  void suggestedIntentTagsKeepsHighestScoreWhenDuplicateTagsReappear() {
    TaskEntry dashboardTask = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskPlanTemplate dashboardTemplate = GridGrindTaskPlanner.templateFor("DASHBOARD");
    List<String> suggestedTags =
        GridGrindGoalPlanner.suggestedIntentTags(
            List.of("dashboard"),
            List.of(
                new GoalPlanReport.Candidate(
                    dashboardTask, 40, List.of("dashboard"), List.of("high"), dashboardTemplate),
                new GoalPlanReport.Candidate(
                    dashboardTask, 20, List.of("dashboard"), List.of("low"), dashboardTemplate)));

    assertEquals(List.copyOf(new java.util.LinkedHashSet<>(suggestedTags)), suggestedTags);
    assertTrue(suggestedTags.contains("dashboard"));
    assertEquals(5, suggestedTags.size());
  }

  @Test
  void suggestedIntentTagsReplacesLowerScoreWhenStrongerDuplicateAppearsLater() {
    TaskEntry dashboardTask = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskPlanTemplate dashboardTemplate = GridGrindTaskPlanner.templateFor("DASHBOARD");
    List<String> suggestedTags =
        GridGrindGoalPlanner.suggestedIntentTags(
            List.of("dashboard"),
            List.of(
                new GoalPlanReport.Candidate(
                    dashboardTask, 20, List.of("dashboard"), List.of("low"), dashboardTemplate),
                new GoalPlanReport.Candidate(
                    dashboardTask, 40, List.of("dashboard"), List.of("high"), dashboardTemplate)));

    assertEquals(List.copyOf(new java.util.LinkedHashSet<>(suggestedTags)), suggestedTags);
    assertTrue(suggestedTags.contains("dashboard"));
    assertEquals(5, suggestedTags.size());
  }

  private static GoalPlanReport.Candidate candidate(
      String taskId, int score, List<String> matchedTerms) {
    return new GoalPlanReport.Candidate(
        GridGrindTaskCatalog.entryFor(taskId).orElseThrow(),
        score,
        matchedTerms,
        List.of("test"),
        GridGrindTaskPlanner.templateFor(taskId));
  }
}
