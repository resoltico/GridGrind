package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/** Verifies that an Error beyond the execution adapter still reaches the canonical CLI boundary. */
class GridGrindCliUnexpectedFailureTest extends GridGrindCliTestSupport {
  @Test
  void outerCliBoundaryConvertsUnexpectedErrorsIntoARejectedCommand() throws Exception {
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, inputs, sink) -> {
              throw new AssertionError("secret execution failure");
            });
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        cli.run(
            stdinExecutionArguments(),
            new ByteArrayInputStream(
                requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                    .getBytes(StandardCharsets.UTF_8)),
            stdout,
            stderr);

    assertEquals(1, exitCode);
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        commandErrorOnStdout(stdout, stderr).primaryProblem().code());
  }
}
