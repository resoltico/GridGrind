package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies that constructor failures retain the request token named by their qualified path. */
class RequestAnalysisLocationTest {
  @Test
  void locatesCompletePlanBindingFailuresAtTheOffendingStepMember() {
    String request =
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
        """;
    byte[] bytes = request.getBytes(StandardCharsets.UTF_8);

    RequestAnalysis analysis = GridGrindJson.analyzeRequest(bytes);

    RequestBindingFailure failure = analysis.bindingFailures().getFirst();
    assertEquals("steps[1].stepId", failure.jsonPath());
    assertEquals(Optional.of(secondIndexOf(bytes, "\"stepId\"")), failure.byteOffset());
    assertTrue(analysis.completePlan().isEmpty());
  }

  @Test
  void returnsPathOnlyWhenAManuallyAssembledAnalysisHasNoRawRequestTree() {
    RequestAnalysis decoded =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    RequestAnalysis manual = new RequestAnalysis(decoded.boundFragments(), List.of());

    assertEquals(
        ProblemContextRequestSurfaces.JsonLocation.pathOnly("source"),
        manual.jsonLocationAt("source"));
    assertTrue(manual.byteOffsetAt("source").isEmpty());
  }

  private static long secondIndexOf(byte[] bytes, String token) {
    String text = new String(bytes, StandardCharsets.UTF_8);
    int firstIndex = text.indexOf(token);
    return text.indexOf(token, firstIndex + token.length());
  }
}
