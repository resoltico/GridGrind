package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

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
import java.util.UUID;
import org.junit.jupiter.api.Test;

/** Argument and failure-classification integration tests for GridGrindCli. */
class GridGrindCliFailureClassificationTest extends GridGrindCliTestSupport {
  @Test
  void versionFlagRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--version", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void helpFlagRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--help", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void printRequestTemplateRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-request-template", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void printExampleRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example", "ASSERTION", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void returnsStructuredJsonErrorForInvalidArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--unknown"}, new ByteArrayInputStream(new byte[0]), stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(2, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("PARSE_ARGUMENTS", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("--unknown"), parseArgumentsContext(failure).argumentName());
    assertEquals("Unknown argument: --unknown", failure.problem().message());
  }

  @Test
  void returnsStructuredJsonErrorWhenArgumentValueIsMissing() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--request"}, new ByteArrayInputStream(new byte[0]), stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(2, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("PARSE_ARGUMENTS", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
    assertEquals("Missing value for --request", failure.problem().message());
  }

  @Test
  void rejectsDuplicateArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", "a.json", "--request", "b.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(2, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    assertEquals(
        "Duplicate argument: --request",
        ((GridGrindResponse.Failure) response).problem().message());
  }

  @Test
  void rejectsDuplicateResponseArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--response", "a.json", "--response", "b.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(2, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    assertEquals(
        "Duplicate argument: --response",
        ((GridGrindResponse.Failure) response).problem().message());
  }

  @Test
  void classifiesInvalidPathArgumentFailures() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--request", "\0"}, new ByteArrayInputStream(new byte[0]), stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(2, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("PARSE_ARGUMENTS", failure.problem().context().stage());
  }

  @Test
  void classifiesIoErrorsWhenRequestFileCannotBeRead() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", "/tmp/does-not-exist.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("READ_REQUEST", failure.problem().context().stage());
  }

  @Test
  void classifiesExecutionErrorsAndWritesFailureToResponsePath() throws IOException {
    Path responsePath =
        Files.createTempDirectory("gridgrind-execution-error-")
            .resolve("nested")
            .resolve("response.json");
    GridGrindCli cli =
        new GridGrindCli(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            new String[] {"--response", responsePath.toString()},
            new ByteArrayInputStream(
                """
                {
                  "source": { "type": "NEW" },
                  "steps": []
                }
                """
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
    GridGrindCli cli =
        new GridGrindCli(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            new String[0],
            new ByteArrayInputStream(
                """
                {
                  "source": { "type": "EXISTING", "path": "/tmp/source.xlsx" },
                  "persistence": { "type": "OVERWRITE" },
                  "steps": []
                }
                """
                    .getBytes(StandardCharsets.UTF_8)),
            stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

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
    GridGrindCli cli =
        new GridGrindCli(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            new String[0],
            new ByteArrayInputStream(
                """
                {
                  "source": { "type": "EXISTING", "path": "/tmp/source.xlsx" },
                  "persistence": { "type": "SAVE_AS", "path": "/tmp/output.xlsx" },
                  "steps": []
                }
                """
                    .getBytes(StandardCharsets.UTF_8)),
            stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

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
        new GridGrindCli(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException();
            });

    try {
      int exitCode =
          cli.run(
              new String[] {"--response", responsePath.toString()},
              new ByteArrayInputStream(
                  """
                    {
                      "source": { "type": "NEW" },
                      "steps": []
                    }
                    """
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

    int exitCode =
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_JSON, failure.problem().code());
    assertEquals("READ_REQUEST", failure.problem().context().stage());
    assertEquals(java.util.Optional.of(1), readRequestContext(failure).jsonLine());
    assertEquals(java.util.Optional.of(2), readRequestContext(failure).jsonColumn());
    assertEquals(java.util.Optional.empty(), readRequestContext(failure).jsonPath());
  }

  @Test
  void classifiesRequestShapeFailuresAsReadRequest() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                {
                  "source": { "type": "NEW" },
                  "steps": [
                    { "stepId": "summary", "target": { "type": "WORKBOOK_CURRENT" } }
                  ]
                }
                """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.problem().code());
    assertEquals("READ_REQUEST", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("steps[0]"), readRequestContext(failure).jsonPath());
    assertEquals(1, failure.problem().causes().size());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.problem().causes().getFirst().code());
    assertFalse(failure.problem().causes().getFirst().message().contains("tools.jackson"));
    assertFalse(failure.problem().causes().getFirst().message().contains("dev.erst.gridgrind"));
  }

  @Test
  void classifiesSemanticRequestValidationAsReadRequest() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                {
                  "source": { "type": "NEW" },
                  "steps": [
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
                }
                """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("READ_REQUEST", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("steps[0].target"), readRequestContext(failure).jsonPath());
    assertEquals(java.util.Optional.of(14), readRequestContext(failure).jsonLine());
    assertTrue(readRequestContext(failure).jsonColumn().isPresent());
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
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", requestPath.toString()},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("READ_REQUEST", failure.problem().context().stage());
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
                  new String[] {"--response", responsePath.toString()},
                  new ByteArrayInputStream(
                      """
                    {
                      "source": { "type": "NEW" },
                      "steps": []
                    }
                    """
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
  void fallsBackToStdoutWhenResponsePathCannotBeWritten() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-response-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--response", responseDirectory.toString()},
                new ByteArrayInputStream(
                    """
                {
                  "source": { "type": "NEW" },
                  "steps": []
                }
                """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("WRITE_RESPONSE", failure.problem().context().stage());
    assertEquals(
        java.util.Optional.of(responseDirectory.toAbsolutePath().toString()),
        writeResponseContext(failure).responsePath());
  }

  @Test
  void preservesOriginalProblemWhenFallbackResponseWriteAlsoFails() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-response-dir-error-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    GridGrindCli cli =
        new GridGrindCli(
            (request, bindings, sink) -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            new String[] {"--response", responseDirectory.toString()},
            new ByteArrayInputStream(
                """
                {
                  "source": { "type": "NEW" },
                  "steps": []
                }
                """
                    .getBytes(StandardCharsets.UTF_8)),
            stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals(2, failure.problem().causes().size());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().causes().get(1).code());
    assertEquals("EXECUTE_REQUEST", failure.problem().causes().get(1).stage());
  }

  @Test
  void classifiesInvalidFormulasAsFormulaErrors() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                {
                  "protocolVersion": "V1",
                  "source": { "type": "NEW" },
                  "persistence": { "type": "NONE" },
                  "execution": {
                    "calculation": { "strategy": { "type": "EVALUATE_ALL" } }
                  },
                  "steps": [
                    { "stepId": "ensure-data", "target": { "type": "SHEET_BY_NAME", "name": "Data" }, "action": { "type": "ENSURE_SHEET" } },
                    { "stepId": "set-formula", "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Data", "address": "A1" }, "action": { "type": "SET_CELL", "value": { "type": "FORMULA", "source": { "type": "INLINE", "text": "SUM(" } } } }
                  ]
                }
                """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("A1"), executeStepContext(failure).address());
    assertEquals(java.util.Optional.of("SUM("), executeStepContext(failure).formula());
  }

  @Test
  void doesNotCloseProvidedStdinWhenReadingRequestFromStandardInput() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    try (TrackingInputStream stdin =
        new TrackingInputStream(
            """
            {
              "source": { "type": "NEW" },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8))) {
      int exitCode = new GridGrindCli().run(new String[0], stdin, stdout);

      assertEquals(0, exitCode);
      assertFalse(stdin.closed());
    }
  }

  @Test
  void rejectsSamePathForRequestAndResponse() throws IOException {
    Path path = Files.createTempFile("gridgrind-same-path-", ".json");

    try {
      ByteArrayOutputStream stdout = new ByteArrayOutputStream();

      int exitCode =
          new GridGrindCli()
              .run(
                  new String[] {"--request", path.toString(), "--response", path.toString()},
                  new ByteArrayInputStream(new byte[0]),
                  stdout);

      GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

      assertEquals(2, exitCode);
      assertInstanceOf(GridGrindResponse.Failure.class, response);
      GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
      assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
      assertEquals("PARSE_ARGUMENTS", failure.problem().context().stage());
      assertEquals(
          "--request and --response must not point to the same path", failure.problem().message());
    } finally {
      Files.deleteIfExists(path);
    }
  }

  @Test
  void rejectsGetWindowWhenCellCountExceedsLimit() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
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
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(
        failure.problem().message().contains("rowCount * columnCount must not exceed"),
        "message should state the limit");
  }
}
