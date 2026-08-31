package dev.erst.gridgrind.engine.api;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers the engine-owned canonical projection for tolerant request-intake findings. */
class GridGrindRequestAnalysisProblemsTest {
  @Test
  void projectsEveryStructuralJsonLocationShape() {
    ProblemContext.ReadRequest byteOnly =
        firstReadRequest(GridGrindJson.analyzeRequest(new byte[0]));
    ProblemContext.ReadRequest pathOnly =
        firstReadRequest(
            GridGrindJson.analyzeRequest(
                "{\"protocolVersion\":\"V2\"}".getBytes(StandardCharsets.UTF_8)));
    ProblemContext.ReadRequest pathAtByteOffset =
        firstReadRequest(GridGrindJson.analyzeRequest("null".getBytes(StandardCharsets.UTF_8)));
    byte[] duplicateRequest =
        "{\"protocolVersion\":\"V2\",\"protocolVersion\":\"V2\"}".getBytes(StandardCharsets.UTF_8);
    ProblemContext.ReadRequest duplicate =
        projected(duplicateRequest).stream()
            .map(GridGrindProblemDetail.Problem::context)
            .filter(ProblemContext.ReadRequest.class::isInstance)
            .map(ProblemContext.ReadRequest.class::cast)
            .filter(context -> context.json().duplicateKeyValue().isPresent())
            .findFirst()
            .orElseThrow();

    assertEquals(ProblemContextRequestSurfaces.JsonLocation.byteOffset(0), byteOnly.json());
    assertEquals(ProblemContextRequestSurfaces.JsonLocation.pathOnly("source"), pathOnly.json());
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset("request", 0),
        pathAtByteOffset.json());
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.duplicateKey(
            "",
            "protocolVersion",
            0,
            new String(duplicateRequest, StandardCharsets.UTF_8)
                .lastIndexOf("\"protocolVersion\"")),
        duplicate.json());
  }

  @Test
  void projectsMalformedSyntaxWithOffsetLineAndColumn() {
    ProblemContext.ReadRequest syntax =
        firstReadRequest(
            GridGrindJson.analyzeRequest("{\n  broken\n}".getBytes(StandardCharsets.UTF_8)));

    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.byteOffsetLineColumn(4, 2, 3), syntax.json());
  }

  @Test
  void projectsLocatedFragmentAndCompletePlanBindingFailures() {
    byte[] fragmentFailure =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "EXISTING", "path": "source.xls" },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """
            .getBytes(StandardCharsets.UTF_8);
    byte[] completePlanFailure =
        """
        {
          "protocolVersion": "V2",
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
        """
            .getBytes(StandardCharsets.UTF_8);

    ProblemContext.BindRequest fragmentContext = firstBindRequest(fragmentFailure);
    ProblemContext.BindRequest completePlanContext = firstBindRequest(completePlanFailure);

    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset(
            "source.path", indexOf(fragmentFailure, "\"path\": \"source.xls\"")),
        fragmentContext.json());
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset(
            "steps[1].stepId", secondIndexOf(completePlanFailure, "\"stepId\"")),
        completePlanContext.json());
  }

  @Test
  void retainsPathOnlyLocationsWhenABindingFailureHasNoSourceOffset() {
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathOnly("steps[0].target"),
        GridGrindRequestAnalysisProblems.locationFor("steps[0].target", Optional.empty()));
  }

  private static ProblemContext.ReadRequest firstReadRequest(
      dev.erst.gridgrind.contract.json.RequestAnalysis analysis) {
    return projected(analysis).stream()
        .map(GridGrindProblemDetail.Problem::context)
        .filter(ProblemContext.ReadRequest.class::isInstance)
        .map(ProblemContext.ReadRequest.class::cast)
        .findFirst()
        .orElseThrow();
  }

  private static ProblemContext.BindRequest firstBindRequest(byte[] request) {
    return projected(request).stream()
        .map(GridGrindProblemDetail.Problem::context)
        .filter(ProblemContext.BindRequest.class::isInstance)
        .map(ProblemContext.BindRequest.class::cast)
        .findFirst()
        .orElseThrow();
  }

  private static List<GridGrindProblemDetail.Problem> projected(byte[] request) {
    return projected(GridGrindJson.analyzeRequest(request));
  }

  private static List<GridGrindProblemDetail.Problem> projected(
      dev.erst.gridgrind.contract.json.RequestAnalysis analysis) {
    return GridGrindRequestAnalysisProblems.project(
        analysis, ProblemContextRequestSurfaces.RequestInput.standardInput());
  }

  private static long indexOf(byte[] bytes, String token) {
    return new String(bytes, StandardCharsets.UTF_8).indexOf(token);
  }

  private static long secondIndexOf(byte[] bytes, String token) {
    String text = new String(bytes, StandardCharsets.UTF_8);
    int firstIndex = text.indexOf(token);
    return text.indexOf(token, firstIndex + token.length());
  }
}
