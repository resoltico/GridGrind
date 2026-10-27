package dev.erst.gridgrind.cli.render;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.discovery.RecipeKeywordMatchReport;
import dev.erst.gridgrind.core.discovery.RecipeType;
import dev.erst.gridgrind.core.discovery.TaskCatalogs;
import dev.erst.gridgrind.core.discovery.TaskEntry;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for machine-readable recipe keyword match reports. */
class RecipeKeywordMatchReportTest {
  @Test
  void reportConstructorCopiesCandidatesAndRequiresExplicitProtocolVersion() {
    TaskEntry task = TaskCatalogs.entryFor("DASHBOARD").orElseThrow();
    RecipeKeywordMatchReport.Candidate candidate =
        new RecipeKeywordMatchReport.Candidate(
            task.id(),
            RecipeType.TASK_STARTER,
            task.narrative().summary(),
            42,
            List.of("dashboard", "chart"),
            List.of("intent tag", "summary"));
    RecipeKeywordMatchReport report =
        new RecipeKeywordMatchReport(
            dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT,
            "monthly dashboard with charts",
            List.of("monthly", "dashboard", "chart"),
            List.of("monthly"),
            List.of("dashboard", "audit"),
            new java.util.ArrayList<>(List.of(candidate)));

    assertEquals(task.id(), report.candidates().getFirst().recipeId());
    assertEquals(
        dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT, report.protocolVersion());
    assertThrows(UnsupportedOperationException.class, () -> report.candidates().add(candidate));

    NullPointerException exception =
        assertThrows(
            NullPointerException.class,
            () ->
                new RecipeKeywordMatchReport(
                    null,
                    "monthly dashboard with charts",
                    List.of("monthly", "dashboard", "chart"),
                    List.of("monthly"),
                    List.of("dashboard", "audit"),
                    List.of(candidate)));
    assertEquals("protocolVersion must not be null", exception.getMessage());
  }

  @Test
  void reportAndCandidateValidationRejectInvalidShapes() {
    TaskEntry task = TaskCatalogs.entryFor("DASHBOARD").orElseThrow();

    IllegalArgumentException blankQuery =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new RecipeKeywordMatchReport(
                    dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT,
                    " ",
                    List.of("dashboard"),
                    List.of(),
                    List.of("dashboard"),
                    List.of()));
    assertEquals("query must not be blank", blankQuery.getMessage());

    IllegalArgumentException duplicateRecipeId =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new RecipeKeywordMatchReport(
                    dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT,
                    "dashboard",
                    List.of("dashboard"),
                    List.of(),
                    List.of("dashboard"),
                    List.of(
                        new RecipeKeywordMatchReport.Candidate(
                            task.id(),
                            RecipeType.TASK_STARTER,
                            task.narrative().summary(),
                            42,
                            List.of("dashboard"),
                            List.of("intent tag")),
                        new RecipeKeywordMatchReport.Candidate(
                            task.id(),
                            RecipeType.TASK_STARTER,
                            task.narrative().summary(),
                            21,
                            List.of("chart"),
                            List.of("capability summary")))));
    assertEquals(
        "candidates must not contain duplicate recipe ids: DASHBOARD",
        duplicateRecipeId.getMessage());

    IllegalArgumentException zeroScore =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new RecipeKeywordMatchReport.Candidate(
                    task.id(),
                    RecipeType.TASK_STARTER,
                    task.narrative().summary(),
                    0,
                    List.of("dashboard"),
                    List.of("intent tag")));
    assertEquals("score must be positive", zeroScore.getMessage());

    IllegalArgumentException emptyMatchedTerms =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new RecipeKeywordMatchReport.Candidate(
                    task.id(),
                    RecipeType.TASK_STARTER,
                    task.narrative().summary(),
                    10,
                    List.of(),
                    List.of("intent tag")));
    assertEquals("matchedTerms must not be empty", emptyMatchedTerms.getMessage());

    IllegalArgumentException emptyMatchSources =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new RecipeKeywordMatchReport.Candidate(
                    task.id(),
                    RecipeType.TASK_STARTER,
                    task.narrative().summary(),
                    10,
                    List.of("dashboard"),
                    List.of()));
    assertEquals("matchSources must not be empty", emptyMatchSources.getMessage());
  }
}
