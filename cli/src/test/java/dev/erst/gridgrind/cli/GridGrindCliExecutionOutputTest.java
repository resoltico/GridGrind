package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

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
}
