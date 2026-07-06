package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused invocation-path tests for stdin discovery and execution behavior. */
class GridGrindCliInvocationTest extends GridGrindCliTestSupport {
  @Test
  void noArgInvocationWithEmptyStandardInputReturnsCliDiagnostic() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        nonInteractiveCli()
            .run(new String[0], new ByteArrayInputStream(new byte[0]), stdout, stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("execute", failure.command());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("No request JSON was provided."));
    assertTrue(
        failure
            .problem()
            .resolution()
            .contains("Standard-input request mode always requires --execution-root"));
  }

  @Test
  void noArgInvocationWithResponsePathWritesCliDiagnosticToFile() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-no-request-response-", ".json");
    Files.deleteIfExists(responsePath);
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
    CliDiagnostic failure = cliDiagnostic(Files.readAllBytes(responsePath));
    CliDiagnostic stderrDiagnostic = cliDiagnosticOnStderr(stderr);
    assertEquals(failure, stderrDiagnostic);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
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

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(
        java.util.Optional.of("--execution-root"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("--execution-root"));
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
  void invocationPinsNonStringSubtypeDiscriminatorsToTheExactTypeField() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            """
                            [
                              {
                                "stepId": "bad-shape",
                                "target": { "type": "WORKBOOK_CURRENT" },
                                "query": { "type": 1 }
                              }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.problem().code());
    assertEquals("execute", failure.command());
    assertEquals(
        java.util.Optional.of("steps[0].query.type"), readRequestContext(failure).jsonPath());
    assertEquals("Field 'type' must be a string", failure.problem().message());
    assertEquals(
        "Replace field 'type' with a JSON string type id.", failure.problem().resolution());
  }

  @Test
  void invocationQualifiesStepTargetMissingTypeMessagesToTheExactNestedPath() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            """
                            [
                              {
                                "stepId": "bad-target",
                                "target": {},
                                "query": { "type": "GET_WORKBOOK_SUMMARY" }
                              }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.problem().code());
    assertEquals(Optional.of("steps[0].target.type"), readRequestContext(failure).jsonPath());
    assertEquals("Missing required field 'steps[0].target.type'", failure.problem().message());
    assertTrue(failure.problem().resolution().contains("steps[0].target.type"));
  }

  @Test
  void executeAndDoctorRequestShareTheSameProblemCoreForMalformedRequests() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-malformed-request-", ".json");
    Files.writeString(
        requestPath,
        minimalRequestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              {
                "stepId": "bad-target",
                "target": {},
                "query": { "type": "GET_WORKBOOK_SUMMARY" }
              }
            ]
            """));
    ByteArrayOutputStream executeStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream executeStderr = new ByteArrayOutputStream();
    ByteArrayOutputStream doctorStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream doctorStderr = new ByteArrayOutputStream();

    int executeExitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", requestPath.toString()},
                InputStream.nullInputStream(),
                executeStdout,
                executeStderr);
    int doctorExitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", requestPath.toString()},
                InputStream.nullInputStream(),
                doctorStdout,
                doctorStderr);

    CliDiagnostic executeDiagnostic = cliDiagnosticOnStderr(executeStdout, executeStderr);
    RequestDoctorReport doctorReport = doctorReport(doctorStdout, doctorStderr);

    assertEquals(1, executeExitCode);
    assertEquals(1, doctorExitCode);
    assertEquals(executeDiagnostic.problem(), doctorReport.primaryProblem().orElseThrow());
    assertEquals(
        Optional.of("steps[0].target.type"), readRequestContext(executeDiagnostic).jsonPath());
    assertEquals(Optional.of("steps[0].target.type"), readRequestContext(doctorReport).jsonPath());
  }

  @Test
  void invocationNamesTheOwnedTargetTypeFieldForCustomTargetShapeFailures() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            """
                            [
                              {
                                "stepId": "bad-target",
                                "target": { "type": 7 },
                                "query": { "type": "GET_WORKBOOK_SUMMARY" }
                              }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.problem().code());
    assertEquals(Optional.of("steps[0].target.type"), readRequestContext(failure).jsonPath());
    assertEquals("Field 'target.type' must be a string", failure.problem().message());
    assertTrue(failure.problem().resolution().contains("target.type"));
  }

  @Test
  void invocationPinsNonXlsxPathViolationsToTheExactPathField() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"SAVE_AS\", \"path\": \"budget.txt\", \"ifExists\": \"REJECT\" }",
                            "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals(Optional.of("persistence.path"), readRequestContext(failure).jsonPath());
    assertEquals("path must end in .xlsx (got: '.txt')", failure.problem().message());
    assertTrue(failure.problem().resolution().contains("persistence.path"));
  }

  @Test
  void noArgInvocationWithInteractiveStandardInputReturnsCliDiagnosticWithoutReadingInput()
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

      CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
      assertEquals(2, exitCode);
      assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
      assertEquals(
          java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
      assertTrue(failure.problem().message().contains("No request JSON was provided."));
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
