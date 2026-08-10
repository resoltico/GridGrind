package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.RequestBindingFailure;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers the public context projection for every request-intake finding. */
class CliRequestAnalysisProblemsTest {
  @Test
  void emptyRequestReportsItsSingleByteZeroSyntaxProblem() {
    List<dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem> problems =
        CliRequestAnalysisProblems.problems(
            GridGrindJson.analyzeRequest(new byte[0]),
            ProblemContextRequestSurfaces.RequestInput.standardInput());

    List<ProblemContext.ReadRequest> contexts =
        problems.stream()
            .map(problem -> assertInstanceOf(ProblemContext.ReadRequest.class, problem.context()))
            .toList();
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.byteOffset(0), contexts.getFirst().json());
    assertEquals(1, contexts.size());
  }

  @Test
  void encodingProblemsRetainTheirByteOffsetWithoutInventingAJsonPath() {
    List<dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem> problems =
        CliRequestAnalysisProblems.problems(
            GridGrindJson.analyzeRequest(new byte[] {(byte) 0xC3, (byte) 0x28}),
            ProblemContextRequestSurfaces.RequestInput.standardInput());

    ProblemContext.ReadRequest context =
        assertInstanceOf(ProblemContext.ReadRequest.class, problems.getFirst().context());
    assertEquals(ProblemContextRequestSurfaces.JsonLocation.byteOffset(0), context.json());
  }

  @Test
  void rootScalarProblemsRetainTheirValueTokenByteOffsetAndRequestPath() {
    List<dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem> problems =
        CliRequestAnalysisProblems.problems(
            GridGrindJson.analyzeRequest("null".getBytes(StandardCharsets.UTF_8)),
            ProblemContextRequestSurfaces.RequestInput.standardInput());

    ProblemContext.ReadRequest context =
        assertInstanceOf(ProblemContext.ReadRequest.class, problems.getFirst().context());
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset("request", 0), context.json());
  }

  @Test
  void unrepresentableScalarsRetainTheirAuthoredPathAndByteOffset() {
    byte[] request =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "zoom",
              "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
              "action": {
                "type": "SET_SHEET_ZOOM",
                "zoomPercent": 999999999999999999999
              }
            }
          ]
        }
        """
            .getBytes(StandardCharsets.UTF_8);

    List<dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem> problems =
        CliRequestAnalysisProblems.problems(
            GridGrindJson.analyzeRequest(request),
            ProblemContextRequestSurfaces.RequestInput.standardInput());

    ProblemContext.ReadRequest context =
        assertInstanceOf(ProblemContext.ReadRequest.class, problems.getFirst().context());
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset(
            "steps[0].action.zoomPercent",
            indexOf(request, "\"zoomPercent\": 999999999999999999999")),
        context.json());
  }

  @Test
  void constructorFailuresUseTheBindingStageAndRetainTheirExactNestedMemberToken() {
    byte[] request =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "EXISTING", "path": "source.xls" },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """
            .getBytes(StandardCharsets.UTF_8);

    List<dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem> problems =
        CliRequestAnalysisProblems.problems(
            GridGrindJson.analyzeRequest(request),
            ProblemContextRequestSurfaces.RequestInput.standardInput());

    ProblemContext.BindRequest context =
        assertInstanceOf(ProblemContext.BindRequest.class, problems.getFirst().context());
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset(
            "source.path", indexOf(request, "\"path\": \"source.xls\"")),
        context.json());
  }

  @Test
  void completePlanFailuresUseTheBindingStageAndRetainTheirExactNestedMemberToken() {
    byte[] request =
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

    List<dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem> problems =
        CliRequestAnalysisProblems.problems(
            GridGrindJson.analyzeRequest(request),
            ProblemContextRequestSurfaces.RequestInput.standardInput());

    ProblemContext.BindRequest context =
        assertInstanceOf(ProblemContext.BindRequest.class, problems.getFirst().context());
    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset(
            "steps[1].stepId", secondIndexOf(request, "\"stepId\"")),
        context.json());
  }

  @Test
  void unlocatedBindingFailuresRemainPathOnlyInsteadOfInventingAByteOffset() {
    RequestBindingFailure failure =
        new RequestBindingFailure(
            new IllegalArgumentException("unlocatable failure"), "request", Optional.empty());

    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathOnly("request"),
        CliRequestAnalysisProblems.locationFor(failure));
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
