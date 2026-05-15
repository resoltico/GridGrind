package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for machine-readable task keyword match reports. */
class TaskKeywordMatchReportTest {
  @Test
  void reportConstructorCopiesCandidatesAndRequiresExplicitProtocolVersion() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskKeywordMatchReport.Candidate candidate =
        new TaskKeywordMatchReport.Candidate(
            task.id(),
            task.narrative().summary(),
            42,
            List.of("dashboard", "chart"),
            List.of("intent tag", "summary"));
    TaskKeywordMatchReport report =
        new TaskKeywordMatchReport(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            "monthly dashboard with charts",
            List.of("monthly", "dashboard", "chart"),
            List.of("monthly"),
            List.of("dashboard", "audit"),
            new java.util.ArrayList<>(List.of(candidate)));

    assertEquals(task.id(), report.candidates().getFirst().taskId());
    assertEquals(
        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
        report.protocolVersion());
    assertThrows(UnsupportedOperationException.class, () -> report.candidates().add(candidate));

    NullPointerException exception =
        assertThrows(
            NullPointerException.class,
            () ->
                new TaskKeywordMatchReport(
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
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();

    IllegalArgumentException blankQuery =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskKeywordMatchReport(
                    dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
                    " ",
                    List.of("dashboard"),
                    List.of(),
                    List.of("dashboard"),
                    List.of()));
    assertEquals("query must not be blank", blankQuery.getMessage());

    IllegalArgumentException duplicateTaskId =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskKeywordMatchReport(
                    dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
                    "dashboard",
                    List.of("dashboard"),
                    List.of(),
                    List.of("dashboard"),
                    List.of(
                        new TaskKeywordMatchReport.Candidate(
                            task.id(),
                            task.narrative().summary(),
                            42,
                            List.of("dashboard"),
                            List.of("intent tag")),
                        new TaskKeywordMatchReport.Candidate(
                            task.id(),
                            task.narrative().summary(),
                            21,
                            List.of("chart"),
                            List.of("capability summary")))));
    assertEquals(
        "candidates must not contain duplicate task ids: DASHBOARD", duplicateTaskId.getMessage());

    IllegalArgumentException zeroScore =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskKeywordMatchReport.Candidate(
                    task.id(),
                    task.narrative().summary(),
                    0,
                    List.of("dashboard"),
                    List.of("intent tag")));
    assertEquals("score must be positive", zeroScore.getMessage());

    IllegalArgumentException emptyMatchedTerms =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskKeywordMatchReport.Candidate(
                    task.id(), task.narrative().summary(), 10, List.of(), List.of("intent tag")));
    assertEquals("matchedTerms must not be empty", emptyMatchedTerms.getMessage());

    IllegalArgumentException emptyMatchSources =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskKeywordMatchReport.Candidate(
                    task.id(), task.narrative().summary(), 10, List.of("dashboard"), List.of()));
    assertEquals("matchSources must not be empty", emptyMatchSources.getMessage());
  }
}
