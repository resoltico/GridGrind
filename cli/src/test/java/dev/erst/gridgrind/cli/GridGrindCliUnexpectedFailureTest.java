package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies that failures after execution begins retain the execution response contract. */
class GridGrindCliUnexpectedFailureTest extends GridGrindCliTestSupport {
  @Test
  void executionStartedErrorsProduceFailedWorkbookResults() throws Exception {
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
    WorkbookResult.Failure failure =
        assertInstanceOf(WorkbookResult.Failure.class, response(stdout, stderr));
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), failure.problem().message());
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR.title(),
        failure.problem().causes().getFirst().message());
    assertFalse(stdout.toString(StandardCharsets.UTF_8).contains("secret execution failure"));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void preExecutionErrorsRemainRejectedCommandErrors(@TempDir Path temporaryDirectory)
      throws Exception {
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, inputs, sink) -> {
              throw new AssertionError("executor must not run");
            },
            () -> {
              throw new AssertionError("secret pre-execution failure");
            });
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    Path executionRoot = temporaryDirectory.resolve("execution-root");

    int exitCode =
        cli.run(
            new String[] {"--doctor-request", "--execution-root", executionRoot.toString()},
            new ByteArrayInputStream(new byte[0]),
            stdout,
            stderr);

    var rejected = commandErrorOnStdout(stdout, stderr);
    assertEquals(1, exitCode);
    assertEquals("REJECTED", rejected.status());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, rejected.primaryProblem().code());
    assertFalse(stdout.toString(StandardCharsets.UTF_8).contains("secret pre-execution failure"));
  }
}
