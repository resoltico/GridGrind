package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.RecipeKeywordMatchReport;
import dev.erst.gridgrind.cli.discovery.RecipeView;
import dev.erst.gridgrind.cli.discovery.TaskArtifactKind;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskGoalKind;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskVerificationKind;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipe;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for English keyword matching over the unified built-in recipe surface. */
class GridGrindRecipeKeywordMatcherTest {
  @Test
  void keywordMatcherRanksDashboardAuditPivotAndBudgetRecipesDeterministically() {
    RecipeKeywordMatchReport dashboardReport =
        GridGrindRecipeKeywordMatcher.reportFor("Create a monthly sales dashboard with charts");
    RecipeKeywordMatchReport auditReport =
        GridGrindRecipeKeywordMatcher.reportFor("Audit an existing workbook for health findings");
    RecipeKeywordMatchReport pivotReport =
        GridGrindRecipeKeywordMatcher.reportFor("build a pivot report from range data");
    RecipeKeywordMatchReport budgetReport =
        GridGrindRecipeKeywordMatcher.reportFor("budget sheet with totals");

    assertEquals("DASHBOARD", dashboardReport.candidates().getFirst().recipeId());
    assertEquals(RecipeView.TASK_STARTER, dashboardReport.candidates().getFirst().view());
    assertTrue(dashboardReport.candidates().getFirst().matchedTerms().contains("dashboard"));
    assertTrue(dashboardReport.candidates().getFirst().matchedTerms().contains("chart"));
    assertTrue(dashboardReport.suggestedIntentTags().contains("dashboard"));
    assertFalse(dashboardReport.candidates().isEmpty());

    assertEquals("AUDIT_EXISTING_WORKBOOK", auditReport.candidates().getFirst().recipeId());
    assertTrue(auditReport.candidates().getFirst().matchedTerms().contains("audit"));
    assertTrue(auditReport.candidates().getFirst().matchedTerms().contains("existing"));

    assertEquals("PIVOT_REPORT", pivotReport.candidates().getFirst().recipeId());
    assertTrue(pivotReport.candidates().getFirst().matchedTerms().contains("pivot"));

    assertEquals("BUDGET", budgetReport.candidates().getFirst().recipeId());
    assertEquals(RecipeView.EXAMPLE, budgetReport.candidates().getFirst().view());
    assertTrue(budgetReport.candidates().getFirst().matchedTerms().contains("budget"));
  }

  @Test
  void keywordMatcherFindsCustomXmlAndWorkbookMaintenanceWorkflows() {
    RecipeKeywordMatchReport customXmlReport =
        GridGrindRecipeKeywordMatcher.reportFor("import custom xml mapping into an existing xlsx");
    RecipeKeywordMatchReport maintenanceReport =
        GridGrindRecipeKeywordMatcher.reportFor(
            "repair broken workbook comments and copy sheets safely");

    assertEquals("CUSTOM_XML_WORKFLOW", customXmlReport.candidates().getFirst().recipeId());
    assertTrue(customXmlReport.candidates().getFirst().matchedTerms().contains("xml"));
    assertTrue(customXmlReport.candidates().getFirst().matchedTerms().contains("mapping"));
    assertTrue(customXmlReport.candidates().getFirst().matchSources().contains("discovery term"));

    assertEquals("WORKBOOK_MAINTENANCE", maintenanceReport.candidates().getFirst().recipeId());
    assertTrue(maintenanceReport.candidates().getFirst().matchedTerms().contains("comment"));
    assertTrue(maintenanceReport.candidates().getFirst().matchedTerms().contains("copy"));
    assertFalse(maintenanceReport.candidates().isEmpty());
  }

  @Test
  void keywordMatcherLeavesUnmatchedVocabularyVisibleAndNormalizesPluralVariants() {
    RecipeKeywordMatchReport noMatchReport =
        GridGrindRecipeKeywordMatcher.reportFor("speaker notes presentation");
    RecipeKeywordMatchReport normalizationReport =
        GridGrindRecipeKeywordMatcher.reportFor(
            "___ galaxies mazes quizzes boxes churches mosses brushes widgets");

    assertTrue(noMatchReport.candidates().isEmpty());
    assertTrue(noMatchReport.unmatchedTerms().contains("speaker"));
    assertEquals(
        GridGrindCliRecipeRegistry.recipes().stream()
            .flatMap(recipe -> recipe.intentTags().stream())
            .distinct()
            .sorted()
            .toList(),
        noMatchReport.suggestedIntentTags());
    assertEquals(
        List.of("galaxy", "maze", "quiz", "box", "church", "moss", "brush", "widget"),
        normalizationReport.normalizedTerms());
    assertTrue(normalizationReport.unmatchedTerms().contains("church"));
    assertTrue(normalizationReport.unmatchedTerms().contains("quiz"));
    assertTrue(normalizationReport.unmatchedTerms().contains("moss"));
    assertTrue(normalizationReport.unmatchedTerms().contains("brush"));
    assertTrue(normalizationReport.candidates().isEmpty());
  }

  @Test
  void candidateScorerUsesWholeFileStemWhenExampleRequestFileNameHasNoJsonSuffix() {
    GridGrindCliRecipe adHocExample =
        new GridGrindCliRecipe(
            "AD_HOC_TEMPLATE",
            RecipeView.EXAMPLE,
            "ad-hoc-template",
            "summary",
            dev.erst.gridgrind.cli.discovery.RecipeAdvisory.SELF_CONTAINED,
            List.of(),
            List.of("template"),
            dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog.requestTemplate());

    RecipeKeywordMatchReport.Candidate candidate =
        GridGrindRecipeKeywordCandidateScorer.candidateFor(
                adHocExample, List.of("ad", "hoc", "template"))
            .orElseThrow();

    assertEquals("AD_HOC_TEMPLATE", candidate.recipeId());
    assertEquals(RecipeView.EXAMPLE, candidate.view());
    assertTrue(candidate.matchedTerms().contains("template"));
  }

  @Test
  void keywordMatcherRejectsNullBlankAndFullyDiscardedQueries() {
    NullPointerException nullQuery =
        assertThrows(
            NullPointerException.class, () -> GridGrindRecipeKeywordMatcher.reportFor(null));
    assertEquals("query must not be null", nullQuery.getMessage());

    IllegalArgumentException blankQuery =
        assertThrows(
            IllegalArgumentException.class, () -> GridGrindRecipeKeywordMatcher.reportFor(" "));
    assertEquals(
        "query must contain at least one searchable term after normalization",
        blankQuery.getMessage());

    IllegalArgumentException discardedQuery =
        assertThrows(
            IllegalArgumentException.class, () -> GridGrindRecipeKeywordMatcher.reportFor("a"));
    assertEquals(
        "query must contain at least one searchable term after normalization",
        discardedQuery.getMessage());
  }

  @Test
  void candidateOrderingBreaksTiesByMatchedTermsThenTaskId() {
    RecipeKeywordMatchReport.Candidate broaderMatch =
        candidate("DASHBOARD", 50, List.of("chart", "dashboard"));
    RecipeKeywordMatchReport.Candidate narrowerMatch =
        candidate("PIVOT_REPORT", 50, List.of("chart"));
    List<RecipeKeywordMatchReport.Candidate> byMatchedTerms =
        new ArrayList<>(List.of(narrowerMatch, broaderMatch));

    byMatchedTerms.sort(GridGrindRecipeKeywordMatcher.candidateOrdering());

    assertEquals(List.of(broaderMatch, narrowerMatch), byMatchedTerms);

    RecipeKeywordMatchReport.Candidate alphabeticallyLater =
        candidate("WORKBOOK_MAINTENANCE", 50, List.of("repair"));
    RecipeKeywordMatchReport.Candidate alphabeticallyEarlier =
        candidate("CUSTOM_XML_WORKFLOW", 50, List.of("repair"));
    List<RecipeKeywordMatchReport.Candidate> byTaskId =
        new ArrayList<>(List.of(alphabeticallyLater, alphabeticallyEarlier));

    byTaskId.sort(GridGrindRecipeKeywordMatcher.candidateOrdering());

    assertEquals(List.of(alphabeticallyEarlier, alphabeticallyLater), byTaskId);
  }

  @Test
  void suggestedIntentTagsKeepsHighestScoreWhenDuplicateTagsReappear() {
    List<String> suggestedTags =
        GridGrindRecipeKeywordMatcher.suggestedIntentTags(
            List.of(
                candidate("DASHBOARD", 40, List.of("dashboard")),
                candidate("DASHBOARD", 20, List.of("dashboard"))));

    assertEquals(List.copyOf(new java.util.LinkedHashSet<>(suggestedTags)), suggestedTags);
    assertTrue(suggestedTags.contains("dashboard"));
    assertFalse(suggestedTags.isEmpty());
  }

  @Test
  void suggestedIntentTagsReplacesLowerScoreWhenStrongerDuplicateAppearsLater() {
    List<String> suggestedTags =
        GridGrindRecipeKeywordMatcher.suggestedIntentTags(
            List.of(
                candidate("DASHBOARD", 20, List.of("dashboard")),
                candidate("DASHBOARD", 40, List.of("dashboard"))));

    assertEquals(List.copyOf(new java.util.LinkedHashSet<>(suggestedTags)), suggestedTags);
    assertTrue(suggestedTags.contains("dashboard"));
    assertFalse(suggestedTags.isEmpty());
  }

  @Test
  void suggestedIntentTagsFallsBackToPublishedIntentVocabularyOnTotalNoMatch() {
    List<String> suggestedTags = GridGrindRecipeKeywordMatcher.suggestedIntentTags(List.of());

    assertEquals(
        GridGrindCliRecipeRegistry.recipes().stream()
            .flatMap(recipe -> recipe.intentTags().stream())
            .distinct()
            .sorted()
            .toList(),
        suggestedTags);
  }

  @Test
  void keywordMatcherProfileAndPhaseSurfaceTextCoversHiddenEnumBranches() {
    assertEquals(
        "overwrite in place",
        GridGrindTaskKeywordSurfaces.persistenceModeSurface(TaskPersistenceMode.OVERWRITE));
    assertEquals(
        "export extract",
        GridGrindTaskKeywordSurfaces.phasePurposeSurface(TaskPhasePurpose.EXPORT));
    assertEquals(
        "binary payload image object file",
        GridGrindTaskKeywordSurfaces.inputKindSurface(TaskInputKind.BINARY_PAYLOAD));
    assertEquals(
        "drawing anchor position placement",
        GridGrindTaskKeywordSurfaces.inputKindSurface(TaskInputKind.DRAWING_ANCHORS));
    assertEquals(
        "new blank create workbook",
        GridGrindTaskKeywordSurfaces.sourceModeSurface(TaskSourceMode.NEW_WORKBOOK));
    assertEquals(
        "author mutate build update",
        GridGrindTaskKeywordSurfaces.mutationModeSurface(TaskMutationMode.MUTATING));
    assertEquals(
        "self contained portable",
        GridGrindTaskKeywordSurfaces.assetModeSurface(TaskAssetMode.SELF_CONTAINED));
    assertEquals(
        "verify validate confirm", GridGrindTaskKeywordSurfaces.goalSurface(TaskGoalKind.VERIFY));
    assertEquals(
        "formula surface formulas",
        GridGrindTaskKeywordSurfaces.artifactSurface(TaskArtifactKind.FORMULA_SURFACE));
    assertEquals(
        "assert assertion invariant checks",
        GridGrindTaskKeywordSurfaces.verificationKindSurface(
            TaskVerificationKind.ASSERTION_CHECKS));
  }

  @Test
  void keywordSurfaceTablesRejectNullKeysWithOwnedMessages() {
    NullPointerException nullGoal =
        assertThrows(
            NullPointerException.class, () -> GridGrindTaskKeywordSurfaces.goalSurface(null));
    assertEquals("goal must not be null", nullGoal.getMessage());
  }

  private static RecipeKeywordMatchReport.Candidate candidate(
      String taskId, int score, List<String> matchedTerms) {
    var task = GridGrindTaskCatalog.entryFor(taskId).orElseThrow();
    return new RecipeKeywordMatchReport.Candidate(
        task.id(),
        RecipeView.TASK_STARTER,
        task.narrative().summary(),
        score,
        matchedTerms,
        List.of("test"));
  }
}
