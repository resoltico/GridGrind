package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for machine-readable task keyword match reports. */
class TaskKeywordMatchReportTest {
  @Test
  void reportConstructorCopiesCandidatesAndRequiresExplicitProtocolVersion() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskPlanTemplate starterTemplate = starterTemplate(task);
    TaskKeywordMatchReport.Candidate candidate =
        new TaskKeywordMatchReport.Candidate(
            task,
            42,
            List.of("dashboard", "chart"),
            List.of("Matched intent tag \"dashboard\" via dashboard."),
            starterTemplate);
    TaskKeywordMatchReport report =
        new TaskKeywordMatchReport(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            "monthly dashboard with charts",
            List.of("monthly", "dashboard", "chart"),
            List.of("monthly"),
            List.of("dashboard", "audit"),
            new java.util.ArrayList<>(List.of(candidate)));

    assertEquals(task.id(), report.candidates().getFirst().task().id());
    assertEquals(
        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
        report.protocolVersion());
    assertThrows(UnsupportedOperationException.class, () -> report.candidates().add(candidate));
    assertEquals(
        "protocolVersion must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new TaskKeywordMatchReport(
                        null,
                        "monthly dashboard with charts",
                        List.of("monthly", "dashboard", "chart"),
                        List.of("monthly"),
                        List.of("dashboard", "audit"),
                        List.of(candidate)))
            .getMessage());
  }

  @Test
  void reportAndCandidateValidationRejectInvalidShapes() {
    TaskEntry task = GridGrindTaskCatalog.entryFor("DASHBOARD").orElseThrow();
    TaskPlanTemplate starterTemplate = starterTemplate(task);

    assertEquals(
        "query must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskKeywordMatchReport(
                        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
                        " ",
                        List.of("dashboard"),
                        List.of(),
                        List.of("dashboard"),
                        List.of()))
            .getMessage());
    assertEquals(
        "candidates must not contain duplicate task ids: DASHBOARD",
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
                                task,
                                42,
                                List.of("dashboard"),
                                List.of("Matched intent tag \"dashboard\" via dashboard."),
                                starterTemplate),
                            new TaskKeywordMatchReport.Candidate(
                                task,
                                21,
                                List.of("chart"),
                                List.of("Matched capability \"set chart\" via chart."),
                                starterTemplate))))
            .getMessage());
    assertEquals(
        "score must be positive",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskKeywordMatchReport.Candidate(
                        task,
                        0,
                        List.of("dashboard"),
                        List.of("Matched intent tag \"dashboard\" via dashboard."),
                        starterTemplate))
            .getMessage());
    assertEquals(
        "starterTemplate task id must match candidate task id",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskKeywordMatchReport.Candidate(
                        task,
                        10,
                        List.of("dashboard"),
                        List.of("Matched intent tag \"dashboard\" via dashboard."),
                        starterTemplate(
                            GridGrindTaskCatalog.entryFor("AUDIT_EXISTING_WORKBOOK")
                                .orElseThrow())))
            .getMessage());
    assertEquals(
        "matchedTerms must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskKeywordMatchReport.Candidate(
                        task,
                        10,
                        List.of(),
                        List.of("Matched intent tag \"dashboard\" via dashboard."),
                        starterTemplate))
            .getMessage());
    assertEquals(
        "reasons must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskKeywordMatchReport.Candidate(
                        task, 10, List.of("dashboard"), List.of(), starterTemplate))
            .getMessage());
  }

  private static TaskPlanTemplate starterTemplate(TaskEntry task) {
    return new TaskPlanTemplate(
        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
        task,
        GridGrindProtocolCatalog.requestTemplate(),
        List.of("note"));
  }
}
