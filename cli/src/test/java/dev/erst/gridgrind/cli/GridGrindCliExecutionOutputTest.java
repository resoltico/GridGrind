package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Focused stdout/stderr integration tests for executed responses. */
class GridGrindCliExecutionOutputTest extends GridGrindCliTestSupport {
  @Test
  void executedFailureResponsesStayOnStdoutWhenNoResponsePathIsConfigured() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            stdinExecutionArguments(),
            new ByteArrayInputStream(
                requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                    .getBytes(StandardCharsets.UTF_8)),
            stdout,
            stderr);

    GridGrindResponse response = response(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void executedAssertionFailuresStayOnStdoutWhenNoResponsePathIsConfigured() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(
                    assertionMismatchRequestJson("actual", "expected")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    assertEquals(1, exitCode);
    assertTrue(stdout.size() > 0, "executed assertion failures must stay on stdout");
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));

    JsonNode response = JsonMapper.builder().build().readTree(stdout.toByteArray());

    assertEquals("ASSERTION_FAILED", response.path("problem").path("code").asText());
    assertEquals("EXECUTE_STEP", response.path("problem").path("context").path("stage").asText());
    assertEquals(
        "EXPECT_CELL_VALUE",
        response.path("problem").path("assertionFailure").path("assertionType").asText());
    assertEquals(
        "assert-a1", response.path("problem").path("assertionFailure").path("stepId").asText());
  }

  private static String assertionMismatchRequestJson(String actualText, String expectedText) {
    return requestJson(
        "{ \"type\": \"NEW\" }",
        "{ \"type\": \"NONE\" }",
        defaultExecutionJson(),
        emptyFormulaEnvironmentJson(),
        """
        [
          { "stepId": "ensure-data", "target": { "type": "SHEET_BY_NAME", "name": "Data" }, "action": { "type": "ENSURE_SHEET" } },
          { "stepId": "set-a1", "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Data", "address": "A1" }, "action": { "type": "SET_CELL", "value": { "type": "TEXT", "source": { "type": "INLINE", "text": "%s" } } } },
          { "stepId": "assert-a1", "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Data", "address": "A1" }, "assertion": { "type": "EXPECT_CELL_VALUE", "expectedValue": { "type": "TEXT", "text": "%s" } } }
        ]
        """
            .formatted(actualText, expectedText));
  }
}
