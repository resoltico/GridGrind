package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCategory;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.PosixFilePermission;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.UUID;
import org.junit.jupiter.api.Test;

/** Argument and diagnostic-classification integration tests for GridGrindCli. */
class GridGrindCliDiagnosticClassificationTest extends GridGrindCliTestSupport {
  @Test
  void versionFlagRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--version", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void helpFlagRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--help", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void printRequestTemplateRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-request-template", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void printRecipeRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-recipe", "--lookup", "ASSERTION", "--request", "ignored.json"
                },
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void returnsStructuredJsonErrorForInvalidArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--unknown"}, new ByteArrayInputStream(new byte[0]), stdout, stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("cli", failure.command());
    assertEquals(java.util.Optional.of("--unknown"), parseArgumentsContext(failure).argumentName());
    assertEquals("Unknown argument: --unknown", failure.problem().message());
    assertEquals(
        List.of("gridgrind --help", "gridgrind --help-protocol", "gridgrind --help-guidance"),
        failure.suggestions());
    assertEquals(
        "Use one exact CLI flag. Start from --help for the synopsis, --help-protocol for the"
            + " grammar, or --help-guidance for workflow-oriented commands.",
        failure.problem().resolution());
  }

  @Test
  void returnsStructuredJsonErrorWhenArgumentValueIsMissing() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--request"}, new ByteArrayInputStream(new byte[0]), stdout, stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("cli", failure.command());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
    assertEquals("Missing value for --request", failure.problem().message());
    assertEquals(
        List.of(
            "gridgrind --request request.json --response response.json",
            "gridgrind --doctor-request --request request.json --response doctor.json"),
        failure.suggestions());
    assertEquals(
        "Provide one readable request JSON file path, or omit --request and pipe one request"
            + " document on standard input.",
        failure.problem().resolution());
  }

  @Test
  void returnsStructuredJsonErrorWhenExecutionRootValueIsMissing() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--execution-root"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("cli", failure.command());
    assertEquals(
        java.util.Optional.of("--execution-root"), parseArgumentsContext(failure).argumentName());
    assertEquals("Missing value for --execution-root", failure.problem().message());
  }

  @Test
  void rejectsDuplicateArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", "a.json", "--request", "b.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals("Duplicate argument: --request", failure.problem().message());
  }

  @Test
  void rejectsDuplicateResponseArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--response", "a.json", "--response", "b.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals("Duplicate argument: --response", failure.problem().message());
  }

  @Test
  void rejectsDuplicateExecutionRootArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--execution-root", "a", "--execution-root", "b"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals("Duplicate argument: --execution-root", failure.problem().message());
  }

  @Test
  void stdinExecutionRequiresExplicitExecutionRoot() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": { "type": "FULL_XSSF" },
                        "journal": { "level": "NORMAL" },
                        "calculation": {
                          "strategy": { "type": "DO_NOT_CALCULATE" },
                          "markRecalculateOnOpen": false
                        }
                      },
                      "formulaEnvironment": {
                        "externalWorkbooks": [],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("execute", failure.command());
    assertEquals(
        java.util.Optional.of("--execution-root"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("--execution-root"));
  }

  @Test
  void classifiesInvalidPathArgumentFailures() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", "\0"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("cli", failure.command());
  }

  @Test
  void classifiesIoErrorsWhenRequestFileCannotBeRead() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    String missingPath = "/tmp/does-not-exist.json";

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", missingPath},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("execute", failure.command());
    assertEquals(Optional.of(missingPath), readRequestContext(failure).requestPath());
    assertEquals("Request file not found: " + Path.of(missingPath), failure.problem().message());
  }

  @Test
  void classifiesIoErrorsWhenRequestPathIsDirectory() throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-request-directory-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    try {
      int exitCode =
          new GridGrindCli()
              .run(
                  new String[] {"--request", requestDirectory.toString()},
                  new ByteArrayInputStream(new byte[0]),
                  stdout,
                  stderr);

      CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
      assertEquals("execute", failure.command());
      assertEquals(
          "Request path is not a regular file: " + requestDirectory.toAbsolutePath().normalize(),
          failure.problem().message());
    } finally {
      Files.deleteIfExists(requestDirectory);
    }
  }

  @Test
  void classifiesIoErrorsWhenRequestFileIsUnreadable() throws IOException {
    Path unreadableFile = Files.createTempFile("gridgrind-unreadable-request-", ".json");
    String requestBody =
        """
        {
          "source": { "type": "NEW" },
          "steps": []
        }
        """;
    Files.writeString(unreadableFile, requestBody, StandardCharsets.UTF_8);
    Set<PosixFilePermission> originalPermissions = Files.getPosixFilePermissions(unreadableFile);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    try {
      Files.setPosixFilePermissions(unreadableFile, Set.of());
      assertFalse(Files.isReadable(unreadableFile), "test setup must make the file unreadable");

      int exitCode =
          new GridGrindCli()
              .run(
                  new String[] {"--request", unreadableFile.toString()},
                  new ByteArrayInputStream(new byte[0]),
                  stdout,
                  stderr);

      CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
      assertEquals("execute", failure.command());
      assertEquals(
          "Request file is not readable: " + unreadableFile.toAbsolutePath().normalize(),
          failure.problem().message());
    } finally {
      Files.setPosixFilePermissions(unreadableFile, originalPermissions);
      Files.deleteIfExists(unreadableFile);
    }
  }

  @Test
  void classifiesExecutionErrorsAndWritesFailureToResponsePath() throws IOException {
    Path responsePath =
        Files.createTempDirectory("gridgrind-execution-error-")
            .resolve("nested")
            .resolve("response.json");
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            stdinExecutionArguments("--response", responsePath.toString()),
            new ByteArrayInputStream(
                requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                    .getBytes(StandardCharsets.UTF_8)),
            new ByteArrayOutputStream());

    GridGrindResponse response = GridGrindJson.readResponse(Files.readAllBytes(responsePath));

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals(GridGrindProblemCategory.INTERNAL, failure.problem().category());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("NEW"), executeRequestContext(failure).sourceType());
    assertEquals(java.util.Optional.of("NONE"), executeRequestContext(failure).persistenceType());
    assertEquals("boom", failure.problem().message());
  }

  @Test
  void classifiesExecutionErrorsWithExistingSourceAndOverwritePersistence() throws IOException {
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
                requestJson(
                        "{ \"type\": \"EXISTING\", \"path\": \"/tmp/source.xlsx\" }",
                        "{ \"type\": \"OVERWRITE\" }",
                        "[]")
                    .getBytes(StandardCharsets.UTF_8)),
            stdout,
            stderr);

    GridGrindResponse response = response(stdout, stderr);

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("EXISTING"), executeRequestContext(failure).sourceType());
    assertEquals(
        java.util.Optional.of("OVERWRITE"), executeRequestContext(failure).persistenceType());
  }

  @Test
  void classifiesExecutionErrorsWithSaveAsPersistence() throws IOException {
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
                requestJson(
                        "{ \"type\": \"EXISTING\", \"path\": \"/tmp/source.xlsx\" }",
                        "{ \"type\": \"SAVE_AS\", \"path\": \"/tmp/output.xlsx\", \"ifExists\": \"REJECT\" }",
                        "[]")
                    .getBytes(StandardCharsets.UTF_8)),
            stdout,
            stderr);

    GridGrindResponse response = response(stdout, stderr);

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("EXISTING"), executeRequestContext(failure).sourceType());
    assertEquals(
        java.util.Optional.of("SAVE_AS"), executeRequestContext(failure).persistenceType());
  }

  @Test
  void fallsBackToExceptionTypeWhenExecutionErrorHasNoMessage() throws IOException {
    Path responsePath = Path.of("gridgrind-cli-response-" + UUID.randomUUID() + ".json");
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException();
            });

    try {
      int exitCode =
          cli.run(
              stdinExecutionArguments("--response", responsePath.toString()),
              new ByteArrayInputStream(
                  requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                      .getBytes(StandardCharsets.UTF_8)),
              new ByteArrayOutputStream());

      GridGrindResponse response = GridGrindJson.readResponse(Files.readAllBytes(responsePath));

      assertEquals(1, exitCode);
      assertInstanceOf(GridGrindResponse.Failure.class, response);
      GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
      assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
      assertEquals("UnsupportedOperationException", failure.problem().message());
    } finally {
      Files.deleteIfExists(responsePath);
    }
  }

  @Test
  void returnsInvalidJsonForMalformedJson() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_JSON, failure.problem().code());
    assertEquals("execute", failure.command());
    assertTrue(readRequestContext(failure).byteOffset().isPresent());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonLine());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonColumn());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonPath());
  }

  @Test
  void classifiesRequestShapeFailuresAsReadRequest() throws IOException {
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
                              { "stepId": "summary", "target": { "type": "WORKBOOK_CURRENT" } }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.problem().code());
    assertEquals("execute", failure.command());
    assertEquals(java.util.Optional.of("steps[0]"), readRequestContext(failure).jsonPath());
    assertFalse(failure.problem().message().contains("tools.jackson"));
    assertFalse(failure.problem().message().contains("dev.erst.gridgrind"));
  }

  @Test
  void classifiesMissingRequiredRootFieldsAsReadRequestShapeFailures() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(
                    """
                    {
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": { "type": "FULL_XSSF" },
                        "journal": { "level": "SUMMARY" },
                        "calculation": {
                          "strategy": { "type": "DO_NOT_CALCULATE" },
                          "markRecalculateOnOpen": false
                        }
                      },
                      "formulaEnvironment": {
                        "externalWorkbooks": [],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.problem().code());
    assertEquals("execute", failure.command());
    assertEquals(Optional.of("protocolVersion"), readRequestContext(failure).jsonPath());
  }

  @Test
  void classifiesSemanticRequestValidationAsReadRequest() throws IOException {
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
                                "stepId": "window",
                                "target": {
                                  "type": "RANGE_RECTANGULAR_WINDOW",
                                  "sheetName": "Budget",
                                  "topLeftAddress": "A1",
                                  "rowCount": 0,
                                  "columnCount": 1
                                },
                                "query": { "type": "GET_WINDOW" }
                              }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("execute", failure.command());
    assertEquals(
        java.util.Optional.of("steps[0].target.rowCount"), readRequestContext(failure).jsonPath());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonLine());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonColumn());
  }

  @Test
  void rejectsOversizedRequestFilesBeforeExecution() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-request-too-large-", ".json");
    Files.writeString(
        requestPath,
        """
        {
          "source": { "type": "NEW" },
          "steps": [],
          "pad": "%s"
        }
        """
            .formatted("x".repeat((int) GridGrindJson.maxRequestDocumentBytes())),
        StandardCharsets.UTF_8);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", requestPath.toString()},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("execute", failure.command());
    assertEquals(
        "Request JSON exceeds the maximum size of 16 MiB (16777216 bytes); move large authored payloads into UTF8_FILE, FILE, or STANDARD_INPUT sources.",
        failure.problem().message());
  }

  @Test
  void writesResponsesToPathsWithoutParentDirectories() throws IOException {
    Path responsePath = Path.of("gridgrind-cli-" + UUID.randomUUID() + ".json");

    try {
      int exitCode =
          new GridGrindCli()
              .run(
                  stdinExecutionArguments("--response", responsePath.toString()),
                  new ByteArrayInputStream(
                      requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                          .getBytes(StandardCharsets.UTF_8)),
                  new ByteArrayOutputStream());

      GridGrindResponse response = GridGrindJson.readResponse(Files.readAllBytes(responsePath));

      assertEquals(0, exitCode);
      assertInstanceOf(GridGrindResponse.Success.class, response);
    } finally {
      Files.deleteIfExists(responsePath);
    }
  }

  @Test
  void fallsBackToStdoutWhenResponsePathCannotBeWrittenAndMirrorsTheProblemOnStderr()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-response-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments("--response", responseDirectory.toString()),
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    GridGrindResponse response = response(stdout, stderr);
    CliDiagnostic stderrDiagnostic = cliDiagnosticOnStderr(stderr);

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(failure.problem(), stderrDiagnostic.problem());
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("WRITE_RESPONSE", failure.problem().context().stage());
    assertEquals(
        java.util.Optional.of(responseDirectory.toAbsolutePath().toString()),
        writeResponseContext(failure).responsePath());
    assertEquals(
        Optional.of("STDOUT"),
        stderrDiagnostic
            .transport()
            .map(
                transport ->
                    switch (transport) {
                      case dev.erst.gridgrind.cli.discovery.CliTransport.StandardOutput _ ->
                          "STDOUT";
                      case dev.erst.gridgrind.cli.discovery.CliTransport.ResponseFile _ -> "FILE";
                    }));
  }

  @Test
  void parseArgumentFailuresWriteStructuredResponsesToResponsePathAndMirrorThemOnStderr()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-parse-error-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--bogus-flag", "--response", responsePath.toString()},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnostic(Files.readAllBytes(responsePath));
    CliDiagnostic stderrDiagnostic = cliDiagnosticOnStderr(stderr);

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(failure, stderrDiagnostic);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("cli", failure.command());
    assertEquals("Unknown argument: --bogus-flag", failure.problem().message());
  }

  @Test
  void preservesOriginalProblemWhenFallbackResponseWriteAlsoFails() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-response-dir-error-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            stdinExecutionArguments("--response", responseDirectory.toString()),
            new ByteArrayInputStream(
                requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                    .getBytes(StandardCharsets.UTF_8)),
            stdout,
            stderr);

    GridGrindResponse response = response(stdout, stderr);
    CliDiagnostic stderrDiagnostic = cliDiagnosticOnStderr(stderr);

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(failure.problem(), stderrDiagnostic.problem());
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals(2, failure.problem().causes().size());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().causes().get(1).code());
    assertEquals("EXECUTE_REQUEST", failure.problem().causes().get(1).stage());
  }

  @Test
  void doesNotCloseProvidedStdinWhenReadingRequestFromStandardInput() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    try (TrackingInputStream stdin =
        new TrackingInputStream(
            requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                .getBytes(StandardCharsets.UTF_8))) {
      int exitCode = new GridGrindCli().run(stdinExecutionArguments(), stdin, stdout);

      assertEquals(0, exitCode);
      assertFalse(stdin.closed());
    }
  }

  @Test
  void rejectsSamePathForRequestAndResponse() throws IOException {
    Path path = Files.createTempFile("gridgrind-same-path-", ".json");

    try {
      ByteArrayOutputStream stdout = new ByteArrayOutputStream();
      ByteArrayOutputStream stderr = new ByteArrayOutputStream();

      int exitCode =
          new GridGrindCli()
              .run(
                  new String[] {"--request", path.toString(), "--response", path.toString()},
                  new ByteArrayInputStream(new byte[0]),
                  stdout,
                  stderr);

      CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

      assertEquals(2, exitCode);
      assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
      assertEquals("cli", failure.command());
      assertEquals(
          "--request and --response must not point to the same path", failure.problem().message());
    } finally {
      Files.deleteIfExists(path);
    }
  }

  @Test
  void rejectsGetWindowWhenCellCountExceedsLimit() throws IOException {
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
                              { "stepId": "ensure-sheet", "target": { "type": "SHEET_BY_NAME", "name": "S" }, "action": { "type": "ENSURE_SHEET" } },
                              {
                                "stepId": "w",
                                "target": {
                                  "type": "RANGE_RECTANGULAR_WINDOW",
                                  "sheetName": "S",
                                  "topLeftAddress": "A1",
                                  "rowCount": 1000,
                                  "columnCount": 1000
                                },
                                "query": { "type": "GET_WINDOW" }
                              }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("execute", failure.command());
    assertTrue(
        failure.problem().message().contains("rowCount * columnCount must not exceed"),
        "message should state the limit");
  }
}
