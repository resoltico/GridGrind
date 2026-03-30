package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.protocol.GridGrindJson;
import dev.erst.gridgrind.protocol.GridGrindProblemCategory;
import dev.erst.gridgrind.protocol.GridGrindProblemCode;
import dev.erst.gridgrind.protocol.GridGrindProtocolCatalog;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.WorkbookReadResult;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.UUID;
import org.junit.jupiter.api.Test;

/** Integration tests for GridGrindCli command-line invocation. */
class GridGrindCliTest {
  @Test
  void readsJsonRequestFromStdinAndWritesJsonResponse() throws IOException {
    String request =
        """
            {
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
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
              "reads": [
                { "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" },
                { "type": "GET_CELLS", "requestId": "cells", "sheetName": "Budget", "addresses": ["A1", "B3"] }
              ]
            }
            """;

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    GridGrindResponse.WorkbookSummary workbook =
        ((WorkbookReadResult.WorkbookSummaryResult) success.reads().get(0)).workbook();
    assertEquals("Budget", workbook.sheetNames().get(0));
    WorkbookReadResult.CellsResult cells = (WorkbookReadResult.CellsResult) success.reads().get(1);
    GridGrindResponse.CellReport.FormulaReport b3Cell =
        (GridGrindResponse.CellReport.FormulaReport) cells.cells().get(1);
    assertEquals("SUM(B2:B2)", b3Cell.formula());
    assertEquals(
        49.0, ((GridGrindResponse.CellReport.NumberReport) b3Cell.evaluation()).numberValue());
  }

  @Test
  void returnsStructuredJsonErrorForInvalidRequest() throws IOException {
    String request =
        """
            {
              "source": { "type": "EXISTING", "path": "/tmp/does-not-exist.xlsx" },
              "persistence": { "type": "NONE" },
              "operations": [],
              "reads": []
            }
            """;

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
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
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "operations": [
                { "type": "ENSURE_SHEET", "sheetName": "Budget" }
              ],
              "reads": [
                { "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" }
              ]
            }
            """);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
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
    assertEquals(
        List.of("Budget"),
        ((WorkbookReadResult.WorkbookSummaryResult) success.reads().getFirst())
            .workbook()
            .sheetNames());
  }

  @Test
  void versionFrom_returnsUnknown_whenImplementationVersionIsAbsent() {
    assertEquals("unknown", GridGrindCli.versionFrom(null));
  }

  @Test
  void versionFrom_returnsVersion_whenImplementationVersionIsPresent() {
    assertEquals("0.4.1", GridGrindCli.versionFrom("0.4.1"));
  }

  @Test
  void versionFlagPrintsVersionLineToStdoutAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--version"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    // Version reads from the JAR manifest Implementation-Version attribute.
    // When running from the test classpath (no JAR), the attribute is absent and "unknown" is used.
    assertEquals("gridgrind unknown\n", stdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void helpFlagsPrintUsageAndReturnExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int longExitCode =
        new GridGrindCli()
            .run(new String[] {"--help"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, longExitCode);
    String help = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(help.contains("GridGrind CLI"));
    assertTrue(help.contains("Usage:"));
    assertTrue(help.contains("Minimal Valid Request:"));
    assertTrue(help.contains("--request <path>"));
    assertTrue(help.contains("--print-request-template"));
    assertTrue(help.contains("--print-protocol-catalog"));
    assertTrue(help.contains("--help, -h"));
    assertTrue(help.contains("blob/main/docs/QUICK_REFERENCE.md"));

    ByteArrayOutputStream shortStdout = new ByteArrayOutputStream();
    int shortExitCode =
        new GridGrindCli()
            .run(new String[] {"-h"}, new ByteArrayInputStream(new byte[0]), shortStdout);
    assertEquals(0, shortExitCode);
    assertEquals(help, shortStdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void helpTextUsesVersionedDocumentationRoutesWhenVersionKnown() {
    String help = GridGrindCli.helpText("0.9.0");

    assertTrue(help.contains("GridGrind CLI 0.9.0"));
    assertTrue(help.contains("ghcr.io/resoltico/gridgrind:0.9.0"));
    assertTrue(help.contains("blob/v0.9.0/docs/QUICK_REFERENCE.md"));
    assertTrue(help.contains("blob/v0.9.0/docs/OPERATIONS.md"));
    assertTrue(help.contains("blob/v0.9.0/docs/ERRORS.md"));
  }

  @Test
  void helpTextContainsProductDescription() {
    String help = GridGrindCli.helpText("1.0.0");

    // The description is either the default fallback or the value from the properties resource.
    // Either way the version line and description line must both be present.
    assertTrue(help.contains("GridGrind CLI 1.0.0"), "help must contain the version line");
    // The description line appears immediately after the version line.
    int versionLineEnd = help.indexOf("GridGrind CLI 1.0.0") + "GridGrind CLI 1.0.0".length();
    String afterVersion = help.substring(versionLineEnd).stripLeading();
    assertFalse(
        afterVersion.startsWith("Usage:"),
        "A description line must appear between the version and Usage:");
  }

  @Test
  void helpTextDockerExampleUsesAbsoluteContainerPaths() {
    String help = GridGrindCli.helpText("1.0.0");

    assertFalse(help.contains("-w /workdir"), "Docker example must not use -w /workdir");
    assertTrue(
        help.contains("/workdir/request.json"), "Docker example must use absolute /workdir/ paths");
    assertTrue(
        help.contains("/workdir/response.json"),
        "Docker example must use absolute /workdir/ paths");
  }

  @Test
  void descriptionFrom_returnsFallback_whenResourceAbsent() {
    // Object.class lives in the bootstrap classloader which has no gridgrind.properties.
    assertEquals("GridGrind", GridGrindCli.descriptionFrom(Object.class));
  }

  @Test
  void descriptionFrom_returnsFallback_whenStreamIsNull() {
    assertEquals("GridGrind", GridGrindCli.descriptionFrom((InputStream) null));
  }

  @Test
  void descriptionFrom_returnsDescription_fromInputStream() {
    InputStream stream =
        new ByteArrayInputStream("description=Custom Description".getBytes(StandardCharsets.UTF_8));
    assertEquals("Custom Description", GridGrindCli.descriptionFrom(stream));
  }

  @Test
  void descriptionFrom_returnsFallback_whenDescriptionIsBlank() {
    InputStream stream =
        new ByteArrayInputStream("description=   ".getBytes(StandardCharsets.UTF_8));
    assertEquals("GridGrind", GridGrindCli.descriptionFrom(stream));
  }

  @Test
  void descriptionFrom_returnsFallback_whenStreamThrowsOnRead() throws IOException {
    try (InputStream broken =
        new InputStream() {
          @Override
          public int read() throws IOException {
            throw new IOException("simulated read failure");
          }
        }) {
      assertEquals("GridGrind", GridGrindCli.descriptionFrom(broken));
    }
  }

  @Test
  void requestTemplateTextWrapsTemplateSerializationFailures() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindCli.requestTemplateText(
                    () -> {
                      throw new IOException("synthetic failure");
                    }));

    assertEquals("Failed to render the built-in request template", failure.getMessage());
    assertEquals("synthetic failure", failure.getCause().getMessage());
  }

  @Test
  void printRequestTemplateFlagPrintsValidRequestAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-request-template"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindRequest request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindProtocolCatalog.requestTemplate(), request);
  }

  @Test
  void printProtocolCatalogFlagPrintsCurrentCatalogAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindProtocolCatalog.Catalog catalog =
        GridGrindJson.readProtocolCatalog(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindProtocolCatalog.catalog(), catalog);
  }

  @Test
  void versionFlagTakesPrecedenceOverOtherFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--version", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    assertEquals(0, exitCode);
    assertEquals("gridgrind unknown\n", stdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void helpFlagTakesPrecedenceOverOtherFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--help", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    assertEquals(0, exitCode);
    assertTrue(stdout.toString(StandardCharsets.UTF_8).contains("GridGrind CLI"));
  }

  @Test
  void printRequestTemplateFlagTakesPrecedenceOverOtherFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-request-template", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindRequest request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindProtocolCatalog.requestTemplate(), request);
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
    assertEquals("--unknown", failure.problem().context().argument());
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
    assertEquals("--request", failure.problem().context().argument());
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
            request -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            new String[] {"--response", responsePath.toString()},
            new ByteArrayInputStream(
                """
                {
                  "source": { "type": "NEW" },
                  "operations": [],
                  "reads": []
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
    GridGrindCli cli =
        new GridGrindCli(
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
                      "source": { "type": "NEW" },
                      "operations": [],
                      "reads": []
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
    assertEquals(1, failure.problem().context().jsonLine());
    assertEquals(2, failure.problem().context().jsonColumn());
    assertNull(failure.problem().context().jsonPath());
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
                  "operations": [],
                  "reads": [
                    { "type": "GET_WORKBOOK_SUMMARY" }
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
    assertEquals("reads[0]", failure.problem().context().jsonPath());
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
                  "operations": [],
                  "reads": [
                    {
                      "type": "GET_WINDOW",
                      "requestId": "window",
                      "sheetName": "Budget",
                      "topLeftAddress": "A1",
                      "rowCount": 0,
                      "columnCount": 1
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
    assertEquals("reads[0]", failure.problem().context().jsonPath());
    assertEquals(12, failure.problem().context().jsonLine());
    assertNotNull(failure.problem().context().jsonColumn());
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
                      "operations": [],
                      "reads": []
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
                  "operations": [],
                  "reads": []
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
    GridGrindCli cli =
        new GridGrindCli(
            request -> {
              throw new UnsupportedOperationException("boom");
            });

    int exitCode =
        cli.run(
            new String[] {"--response", responseDirectory.toString()},
            new ByteArrayInputStream(
                """
                {
                  "source": { "type": "NEW" },
                  "operations": [],
                  "reads": []
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
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                {
                  "protocolVersion": "V1",
                  "source": { "type": "NEW" },
                  "persistence": { "type": "NONE" },
                  "operations": [
                    { "type": "ENSURE_SHEET", "sheetName": "Data" },
                    { "type": "SET_CELL", "sheetName": "Data", "address": "A1", "value": { "type": "FORMULA", "formula": "SUM(" } },
                    { "type": "EVALUATE_FORMULAS" }
                  ],
                  "reads": []
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
              "source": { "type": "NEW" },
              "operations": [],
              "reads": []
            }
            """
                .getBytes(StandardCharsets.UTF_8))) {
      int exitCode = new GridGrindCli().run(new String[0], stdin, stdout);

      assertEquals(0, exitCode);
      assertFalse(stdin.closed);
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
                      "operations": [ { "type": "ENSURE_SHEET", "sheetName": "S" } ],
                      "reads": [
                        {
                          "type": "GET_WINDOW",
                          "requestId": "w",
                          "sheetName": "S",
                          "topLeftAddress": "A1",
                          "rowCount": 1000,
                          "columnCount": 1000
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

  @Test
  void helpTextMentionsZeroSheetsForNewWorkbook() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(
        help.contains("zero sheets"),
        "help must mention that a NEW workbook starts with zero sheets");
    assertTrue(help.contains("ENSURE_SHEET"), "help must mention ENSURE_SHEET as the remedy");
  }

  @Test
  void helpTextListsKeyLimitsUpfront() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(help.contains(".xlsx"), "help must state the .xlsx-only file format limit");
    assertTrue(help.contains("31"), "help must state the 31-character sheet name limit");
    assertTrue(
        help.contains("250,000"), "help must state the GET_WINDOW / GET_SHEET_SCHEMA cell limit");
    assertTrue(help.contains("255"), "help must state the column width limit");
    assertTrue(help.contains("1638"), "help must state the row height limit");
    assertTrue(
        help.contains("NUMBER"),
        "help must note that DATE/DATE_TIME inputs are stored as NUMBER on read-back");
  }

  @Test
  void printProtocolCatalogWithNonOperationTrailingArgReturnsFullCatalog() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--version"},
                InputStream.nullInputStream(),
                stdout);

    // --version after --print-protocol-catalog is not --operation so filter stays null;
    // full catalog is returned (--version is silently ignored as a trailing arg)
    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("operationTypes"), "full catalog must contain operationTypes key");
  }

  @Test
  void printProtocolCatalogWithOperationFilterReturnsMatchingEntry() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--operation", "SET_CELL"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"SET_CELL\""), "output must contain the entry id");
  }

  @Test
  void printProtocolCatalogWithUnknownOperationReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--operation", "BOGUS_XYZ"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(2, exitCode);
    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("BOGUS_XYZ"));
  }

  @Test
  void helpTextMentionsOptionalRequestFields() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(
        help.contains("protocolVersion"), "help must mention that protocolVersion is optional");
    assertTrue(
        help.contains("persistence is optional"), "help must mention that persistence is optional");
    assertTrue(
        help.contains("operations is optional"), "help must mention that operations is optional");
    assertTrue(help.contains("reads is optional"), "help must mention that reads is optional");
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
