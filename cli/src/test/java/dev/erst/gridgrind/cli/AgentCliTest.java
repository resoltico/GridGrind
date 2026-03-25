package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.protocol.GridGrindJson;
import dev.erst.gridgrind.protocol.GridGrindProblemCategory;
import dev.erst.gridgrind.protocol.GridGrindProblemCode;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.UUID;
import org.junit.jupiter.api.Test;

/** Integration tests for AgentCli command-line invocation. */
class AgentCliTest {
  @Test
  void readsJsonRequestFromStdinAndWritesJsonResponse() throws IOException {
    String request =
        """
            {
              "source": { "mode": "NEW" },
              "persistence": { "mode": "NONE" },
              "operations": [
                { "type": "ENSURE_SHEET", "sheetName": "Budget" },
                { "type": "APPEND_ROW", "sheetName": "Budget", "values": [
                  { "type": "TEXT", "text": "Item" },
                  { "type": "TEXT", "text": "Amount" }
                ] },
                { "type": "APPEND_ROW", "sheetName": "Budget", "values": [
                  { "type": "TEXT", "text": "Hosting" },
                  { "type": "NUMBER", "number": 49.0 }
                ] },
                { "type": "SET_CELL", "sheetName": "Budget", "address": "B3", "value": { "type": "FORMULA", "formula": "SUM(B2:B2)" } },
                { "type": "EVALUATE_FORMULAS" }
              ],
              "analysis": {
                "sheets": [
                  { "sheetName": "Budget", "cells": ["A1", "B3"], "previewRowCount": 3, "previewColumnCount": 2 }
                ]
              }
            }
            """;

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new AgentCli()
            .run(
                new String[0],
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals("Budget", success.workbook().sheetNames().get(0));
    GridGrindResponse.CellReport.FormulaReport b3Cell =
        (GridGrindResponse.CellReport.FormulaReport)
            success.sheets().get(0).requestedCells().get(1);
    assertEquals("SUM(B2:B2)", b3Cell.formula());
    assertEquals(
        49.0, ((GridGrindResponse.CellReport.NumberReport) b3Cell.evaluation()).numberValue());
  }

  @Test
  void returnsStructuredJsonErrorForInvalidRequest() throws IOException {
    String request =
        """
            {
              "source": { "mode": "EXISTING_FILE", "path": "/tmp/does-not-exist.xlsx" },
              "persistence": { "mode": "NONE" },
              "operations": [],
              "analysis": { "sheets": [] }
            }
            """;

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new AgentCli()
            .run(
                new String[0],
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertTrue(failure.problem().message().contains("Workbook does not exist"));
  }

  @Test
  void readsJsonRequestFromFileAndWritesJsonResponseToFile() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-request-", ".json");
    Path responsePath =
        Files.createTempDirectory("gridgrind-response-").resolve("nested").resolve("response.json");

    Files.writeString(
        requestPath,
        """
            {
              "source": { "mode": "NEW" },
              "persistence": { "mode": "NONE" },
              "operations": [
                { "type": "ENSURE_SHEET", "sheetName": "Budget" }
              ],
              "analysis": { "sheets": [] }
            }
            """);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new AgentCli()
            .run(
                new String[] {
                  "--request", requestPath.toString(), "--response", responsePath.toString()
                },
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(Files.readAllBytes(responsePath));

    assertEquals(0, exitCode);
    assertEquals(0, stdout.size());
    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals(List.of("Budget"), success.workbook().sheetNames());
  }

  @Test
  void returnsStructuredJsonErrorForInvalidArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new AgentCli()
            .run(new String[] {"--unknown"}, new ByteArrayInputStream(new byte[0]), stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(2, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("PARSE_ARGUMENTS", failure.problem().context().stage());
    assertEquals("--unknown", failure.problem().context().argument());
    assertEquals("Unknown argument: --unknown", failure.problem().message());
  }

  @Test
  void returnsStructuredJsonErrorWhenArgumentValueIsMissing() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new AgentCli()
            .run(new String[] {"--request"}, new ByteArrayInputStream(new byte[0]), stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(2, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("PARSE_ARGUMENTS", failure.problem().context().stage());
    assertEquals("--request", failure.problem().context().argument());
    assertEquals("Missing value for --request", failure.problem().message());
  }

  @Test
  void rejectsDuplicateArguments() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new AgentCli()
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
        new AgentCli()
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
        new AgentCli()
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
        new AgentCli()
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
    AgentCli cli =
        new AgentCli(
            request -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            new String[] {"--response", responsePath.toString()},
            new ByteArrayInputStream(
                """
                {
                  "source": { "mode": "NEW" },
                  "operations": [],
                  "analysis": { "sheets": [] }
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
    assertEquals("boom", failure.problem().message());
  }

  @Test
  void fallsBackToExceptionTypeWhenExecutionErrorHasNoMessage() throws IOException {
    Path responsePath = Path.of("gridgrind-cli-response-" + UUID.randomUUID() + ".json");
    AgentCli cli =
        new AgentCli(
            request -> {
              throw new UnsupportedOperationException();
            });

    try {
      int exitCode =
          cli.run(
              new String[] {"--response", responsePath.toString()},
              new ByteArrayInputStream(
                  """
                    {
                      "source": { "mode": "NEW" },
                      "operations": [],
                      "analysis": { "sheets": [] }
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
  void returnsInvalidRequestForMalformedJson() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new AgentCli()
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
    assertEquals(1, failure.problem().context().jsonLine());
    assertEquals(2, failure.problem().context().jsonColumn());
    assertNull(failure.problem().context().jsonPath());
  }

  @Test
  void classifiesSemanticRequestValidationAsReadRequest() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new AgentCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                {
                  "source": { "mode": "NEW" },
                  "operations": [],
                  "analysis": {
                    "sheets": [
                      { "sheetName": "Budget", "previewRowCount": 0, "previewColumnCount": 1 }
                    ]
                  }
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
    assertEquals("analysis.sheets[0]", failure.problem().context().jsonPath());
    assertEquals(6, failure.problem().context().jsonLine());
    assertNotNull(failure.problem().context().jsonColumn());
  }

  @Test
  void writesResponsesToPathsWithoutParentDirectories() throws IOException {
    Path responsePath = Path.of("gridgrind-cli-" + UUID.randomUUID() + ".json");

    try {
      int exitCode =
          new AgentCli()
              .run(
                  new String[] {"--response", responsePath.toString()},
                  new ByteArrayInputStream(
                      """
                    {
                      "source": { "mode": "NEW" },
                      "operations": [],
                      "analysis": { "sheets": [] }
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
        new AgentCli()
            .run(
                new String[] {"--response", responseDirectory.toString()},
                new ByteArrayInputStream(
                    """
                {
                  "source": { "mode": "NEW" },
                  "operations": [],
                  "analysis": { "sheets": [] }
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
        responseDirectory.toAbsolutePath().toString(), failure.problem().context().responsePath());
  }

  @Test
  void preservesOriginalProblemWhenFallbackResponseWriteAlsoFails() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-response-dir-error-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    AgentCli cli =
        new AgentCli(
            request -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            new String[] {"--response", responseDirectory.toString()},
            new ByteArrayInputStream(
                """
                {
                  "source": { "mode": "NEW" },
                  "operations": [],
                  "analysis": { "sheets": [] }
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
    assertEquals("GridGrindProblem", failure.problem().causes().get(1).type());
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR.name(), failure.problem().causes().get(1).className());
  }

  @Test
  void classifiesInvalidFormulasAsFormulaErrors() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new AgentCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                {
                  "protocolVersion": "V1",
                  "source": { "mode": "NEW" },
                  "persistence": { "mode": "NONE" },
                  "operations": [
                    { "type": "SET_CELL", "sheetName": "Data", "address": "A1", "value": { "type": "FORMULA", "formula": "SUM(" } },
                    { "type": "EVALUATE_FORMULAS" }
                  ],
                  "analysis": { "sheets": [] }
                }
                """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("A1", failure.problem().context().address());
    assertEquals("SUM(", failure.problem().context().formula());
  }

  @Test
  void doesNotCloseProvidedStdinWhenReadingRequestFromStandardInput() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    try (TrackingInputStream stdin =
        new TrackingInputStream(
            """
            {
              "source": { "mode": "NEW" },
              "operations": [],
              "analysis": { "sheets": [] }
            }
            """
                .getBytes(StandardCharsets.UTF_8))) {
      int exitCode = new AgentCli().run(new String[0], stdin, stdout);

      assertEquals(0, exitCode);
      assertFalse(stdin.closed);
    }
  }

  /** ByteArrayInputStream that records whether {@code close()} was called. */
  private static final class TrackingInputStream extends ByteArrayInputStream {
    private boolean closed;

    private TrackingInputStream(byte[] bytes) {
      super(bytes);
    }

    @Override
    public void close() throws IOException {
      closed = true;
      super.close();
    }
  }
}
