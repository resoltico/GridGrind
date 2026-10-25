package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.RequestDiagnosticRedactor;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies that constructor failures retain the request token named by their qualified path. */
class RequestIntakeLocationTest {
  @Test
  void locatesCompletePlanBindingFailuresAtTheOffendingStepMember() {
    String request =
        """
        {
          "protocolVersion": "V3",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "duplicate",
              "target": { "type": "WORKBOOK_CURRENT" },
              "query": { "type": "GET_WORKBOOK_SUMMARY" }
            },
            {
              "stepId": "duplicate",
              "target": { "type": "WORKBOOK_CURRENT" },
              "query": { "type": "GET_WORKBOOK_SUMMARY" }
            }
          ]
        }
        """;
    byte[] bytes = request.getBytes(StandardCharsets.UTF_8);

    RequestIntake analysis = GridGrindJson.analyzeRequest(bytes);

    RequestBindingFailure failure = analysis.bindingFailures().getFirst();
    assertEquals("steps[1].stepId", failure.jsonPath());
    assertEquals(Optional.of(secondIndexOf(bytes, "\"stepId\"")), failure.byteOffset());
    assertTrue(analysis.completePlan().isEmpty());
  }

  @Test
  void returnsPathOnlyWhenAManuallyAssembledAnalysisHasNoRawRequestTree() {
    RequestIntake decoded =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V3",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    RequestIntake manual = manualIntake(decoded.boundFragments(), List.of());

    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathOnly("source"),
        manual.jsonLocationAt("source"));
    assertTrue(manual.byteOffsetAt("source").isEmpty());
  }

  @Test
  void locatesNestedCompactCellsAndCompositeAssertionsAtTheirAuthoredElements() {
    RequestIntake gridAnalysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V3",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [{
                "stepId": "set-range",
                "target": { "type": "RANGE_BY_RANGE", "sheetName": "Budget", "range": "A1" },
                "action": { "type": "SET_RANGE", "rows": { "type": "NUMBER", "cells": [[1e999]] } }
              }]
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    RequestIntake assertionAnalysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V3",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [{
                "stepId": "all-of",
                "target": { "type": "WORKBOOK_CURRENT" },
                "assertion": { "type": "ALL_OF", "assertions": [] }
              }]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    assertEquals(
        "steps[0].action.rows.cells[0][0]",
        gridAnalysis.structuralProblems().getFirst().jsonPath().orElseThrow());
    assertTrue(gridAnalysis.structuralProblems().getFirst().byteOffset().isPresent());
    assertEquals(
        "steps[0].assertion.assertions",
        assertionAnalysis.bindingFailures().getFirst().jsonPath(),
        () -> assertionAnalysis.bindingFailures().getFirst().exception().toString());
    assertTrue(assertionAnalysis.bindingFailures().getFirst().byteOffset().isPresent());
  }

  private static long secondIndexOf(byte[] bytes, String token) {
    String text = new String(bytes, StandardCharsets.UTF_8);
    int firstIndex = text.indexOf(token);
    return text.indexOf(token, firstIndex + token.length());
  }

  private static RequestIntake manualIntake(
      dev.erst.gridgrind.core.problem.RequestBoundFragments fragments,
      java.util.List<dev.erst.gridgrind.core.problem.RequestStructuralProblem> problems) {
    return RequestIntakeBinder.bind(
        fragments,
        problems,
        java.util.List.of(),
        RequestDiagnosticRedactor.forRequestType(Object.class),
        java.util.Optional.empty(),
        java.util.List.of());
  }
}
