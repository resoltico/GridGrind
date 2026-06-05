package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused invocation-path tests for stdin discovery and execution behavior. */
class GridGrindCliInvocationTest extends GridGrindCliTestSupport {
  @Test
  void noArgInvocationWithEmptyStandardInputReturnsCliFailure() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        nonInteractiveCli()
            .run(new String[0], new ByteArrayInputStream(new byte[0]), stdout, stderr);

    CliFailureReport failure = cliFailureOnStdout(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals("execute", failure.command());
    assertEquals(java.util.Optional.of("--request"), failure.argument());
    assertTrue(failure.message().contains("No request JSON was provided."));
    assertTrue(
        failure.resolution().orElseThrow().contains("bare gridgrind invocation now expects"));
  }

  @Test
  void noArgInvocationWithResponsePathWritesCliFailureToFile() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-no-request-response-", ".json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        nonInteractiveCli()
            .run(
                new String[] {"--response", responsePath.toString()},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    CliFailureReport failure = cliFailure(Files.readAllBytes(responsePath));
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--request"), failure.argument());
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains("GridGrind wrote the CLI failure report to"));
  }

  @Test
  void noArgInvocationRequiresExecutionRootWhenStandardInputContainsARequest() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        nonInteractiveCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliFailureReport failure = cliFailureOnStdout(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--execution-root"), failure.argument());
    assertTrue(failure.message().contains("--execution-root"));
  }

  @Test
  void noArgInvocationExecutesWhenStandardInputContainsARequestAndExecutionRootIsExplicit()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    Path workspace = Files.createTempDirectory("gridgrind-no-arg-root-");

    int exitCode =
        nonInteractiveCli()
            .run(
                new String[] {"--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    assertEquals(0, exitCode);
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
    assertInstanceOf(
        GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
  }

  @Test
  void noArgInvocationWithInteractiveStandardInputReturnsCliFailureWithoutReadingInput()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    try (InputStream blockingStdin =
        new InputStream() {
          @Override
          public int read() throws IOException {
            throw new AssertionError("interactive no-arg execution must not read stdin");
          }

          @Override
          public int read(byte[] b, int off, int len) throws IOException {
            throw new AssertionError("interactive no-arg execution must not read stdin");
          }
        }) {
      int exitCode = interactiveCli().run(new String[0], blockingStdin, stdout, stderr);

      CliFailureReport failure = cliFailureOnStdout(stdout, stderr);
      assertEquals(2, exitCode);
      assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
      assertEquals(java.util.Optional.of("--request"), failure.argument());
      assertTrue(failure.message().contains("No request JSON was provided."));
    }
  }

  private static GridGrindCli nonInteractiveCli() {
    return GridGrindCli.forTesting(
        (ignoredRequest, ignoredBindings, ignoredSink) ->
            GridGrindResponses.success(List.of(), List.of(), List.of()),
        () -> false);
  }

  private static GridGrindCli interactiveCli() {
    return GridGrindCli.forTesting(
        (ignoredRequest, ignoredBindings, ignoredSink) -> {
          throw new AssertionError("interactive no-arg invocation must not execute a request");
        },
        () -> true);
  }
}
