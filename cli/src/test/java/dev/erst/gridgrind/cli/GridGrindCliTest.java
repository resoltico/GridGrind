package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.engine.api.GridGrindJournalSink;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

/** Execution and transport integration tests for GridGrindCli command-line invocation. */
class GridGrindCliTest extends GridGrindCliTestSupport {
  /**
   * Keeps the historical inner-class name stable so incremental test builds do not retain a stale
   * {@code GridGrindCliTest$TrackingInputStream.class} from before the support extraction.
   */
  @SuppressWarnings("unused")
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

    boolean closed() {
      return closed;
    }
  }

  protected static String[] stdinExecutionArguments(String... trailingArguments)
      throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-cli-stdin-");
    String[] args = new String[2 + trailingArguments.length];
    args[0] = "--execution-root";
    args[1] = workspace.toString();
    System.arraycopy(trailingArguments, 0, args, 2, trailingArguments.length);
    return args;
  }

  @Test
  void cliJournalWriterReturnsNoopWhenRequestIsMissing() {
    CliJournalWriter writer = new CliJournalWriter();

    assertSame(GridGrindJournalSink.NOOP, writer.sinkFor(null, OutputStream.nullOutputStream()));
  }

  @Test
  void cliJournalWriterSwallowsBestEffortIoFailures() throws IOException {
    WorkbookPlan request =
        GridGrindJson.readRequest(
            requestJson(
                    "{ \"type\": \"NEW\" }",
                    "{ \"type\": \"NONE\" }",
                    verboseExecutionJson(),
                    emptyFormulaEnvironmentJson(),
                    "[]")
                .getBytes(StandardCharsets.UTF_8));
    CliJournalWriter writer = new CliJournalWriter();
    try (OutputStream broken =
        new OutputStream() {
          @Override
          public void write(int b) throws IOException {
            throw new IOException("boom");
          }
        }) {
      assertDoesNotThrow(
          () ->
              writer
                  .sinkFor(request, broken)
                  .emit(
                      new ExecutionJournal.Event(
                          "2026-04-18T11:45:00Z",
                          "OPEN",
                          "opened",
                          Optional.empty(),
                          Optional.empty())));
    }
  }

  @Test
  void cliJournalWriterIncludesStepMetadataWhenPresent() throws IOException {
    WorkbookPlan request =
        GridGrindJson.readRequest(
            requestJson(
                    "{ \"type\": \"NEW\" }",
                    "{ \"type\": \"NONE\" }",
                    verboseExecutionJson(),
                    emptyFormulaEnvironmentJson(),
                    "[]")
                .getBytes(StandardCharsets.UTF_8));
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    new CliJournalWriter()
        .sinkFor(request, stderr)
        .emit(
            new ExecutionJournal.Event(
                "2026-04-18T11:45:00Z",
                "STEP",
                "wrote cell",
                Optional.of(7),
                Optional.of("step-007")));

    assertEquals(
        "[gridgrind] 2026-04-18T11:45:00Z STEP stepId=step-007 stepIndex=7 wrote cell"
            + System.lineSeparator(),
        stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void readsJsonRequestFromStdinAndWritesJsonResponse() throws IOException {
    String request =
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            evaluateAllExecutionJson(),
            emptyFormulaEnvironmentJson(),
            """
            [
              { "stepId": "ensure-budget", "target": { "type": "SHEET_BY_NAME", "name": "Budget" }, "action": { "type": "ENSURE_SHEET" } },
              { "stepId": "append-header", "target": { "type": "SHEET_BY_NAME", "name": "Budget" }, "action": { "type": "APPEND_ROW", "values": { "type": "TYPED", "cells": [
                { "type": "TEXT", "source": { "type": "INLINE", "text": "Item" } },
                { "type": "TEXT", "source": { "type": "INLINE", "text": "Amount" } }
              ] } } },
              { "stepId": "append-hosting", "target": { "type": "SHEET_BY_NAME", "name": "Budget" }, "action": { "type": "APPEND_ROW", "values": { "type": "TYPED", "cells": [
                { "type": "TEXT", "source": { "type": "INLINE", "text": "Hosting" } },
                { "type": "NUMBER", "number": 49.0 }
              ] } } },
              { "stepId": "set-total", "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "B3" }, "action": { "type": "SET_CELL", "value": { "type": "FORMULA", "source": { "type": "INLINE", "text": "SUM(B2:B2)" } } } },
              { "stepId": "workbook", "target": { "type": "WORKBOOK_CURRENT" }, "query": { "type": "GET_WORKBOOK_SUMMARY" } },
              { "stepId": "cells", "target": { "type": "CELL_BY_ADDRESSES", "sheetName": "Budget", "addresses": ["A1", "B3"] }, "query": { "type": "GET_CELLS", "projection": { "facets": ["VALUE", "FORMULA"] } } }
            ]
            """);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals(List.of(), success.warnings());
    WorkbookSummary workbook =
        ((WorkbookInspectionResult.WorkbookSummaryResult) success.inspections().get(0)).workbook();
    assertEquals("Budget", workbook.sheetNames().get(0));
    SheetInspectionResult.CellsResult cells =
        (SheetInspectionResult.CellsResult) success.inspections().get(1);
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport b3Cell =
        (dev.erst.gridgrind.contract.dto.CellReport.FormulaReport) cells.cells().get(1);
    assertEquals("SUM(B2:B2)", b3Cell.formula().orElseThrow());
    assertEquals(
        49.0,
        assertInstanceOf(CellValueReport.NumberValue.class, b3Cell.evaluation().orElseThrow())
            .numberValue());
  }

  @Test
  void executionResponsesAreCompactByDefaultAndIndentedWithPretty() throws IOException {
    String request = requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]");

    ByteArrayOutputStream compactStdout = new ByteArrayOutputStream();
    int compactExitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                compactStdout);

    ByteArrayOutputStream prettyStdout = new ByteArrayOutputStream();
    int prettyExitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments("--pretty"),
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                prettyStdout);

    assertEquals(0, compactExitCode);
    assertEquals(0, prettyExitCode);
    assertInstanceOf(
        GridGrindResponse.Success.class, GridGrindJson.readResponse(compactStdout.toByteArray()));
    assertInstanceOf(
        GridGrindResponse.Success.class, GridGrindJson.readResponse(prettyStdout.toByteArray()));
    assertEquals(1L, compactStdout.toString(StandardCharsets.UTF_8).lines().count());
    assertTrue(prettyStdout.toString(StandardCharsets.UTF_8).startsWith("{\n"));
    assertTrue(prettyStdout.toString(StandardCharsets.UTF_8).contains("\n  \"status\" : "));
  }

  @Test
  void rejectsStandardInputAuthoredValuesWhenRequestAlsoUsesStdin() throws IOException {
    String request =
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              { "stepId": "ensure-budget", "target": { "type": "SHEET_BY_NAME", "name": "Budget" }, "action": { "type": "ENSURE_SHEET" } },
              {
                "stepId": "set-title",
                "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                "action": {
                  "type": "SET_CELL",
                  "value": {
                    "type": "TEXT",
                    "source": { "type": "STANDARD_INPUT" }
                  }
                }
              }
            ]
            """);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals("execute", failure.command());
    assertEquals(java.util.Optional.of("--request"), failure.argument());
    assertTrue(failure.message().contains("STANDARD_INPUT"));
  }

  @Test
  void bindsStandardInputToSourceBackedValuesWhenRequestComesFromFile() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-stdin-request-", ".json");
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              { "stepId": "ensure-budget", "target": { "type": "SHEET_BY_NAME", "name": "Budget" }, "action": { "type": "ENSURE_SHEET" } },
              {
                "stepId": "set-title",
                "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                "action": {
                  "type": "SET_CELL",
                  "value": {
                    "type": "TEXT",
                    "source": { "type": "STANDARD_INPUT" }
                  }
                }
              },
              {
                "stepId": "cells",
                "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                "query": { "type": "GET_CELLS" }
              }
            ]
            """),
        StandardCharsets.UTF_8);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--request", requestPath.toString()},
                new ByteArrayInputStream("Quarterly Budget".getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse.Success success =
        assertInstanceOf(
            GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
    SheetInspectionResult.CellsResult cells =
        assertInstanceOf(SheetInspectionResult.CellsResult.class, success.inspections().getFirst());
    dev.erst.gridgrind.contract.dto.CellReport.TextReport a1 =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().getFirst());
    assertEquals(0, exitCode);
    assertEquals("Quarterly Budget", a1.textValue().orElseThrow());
  }

  @Test
  void passesExplicitTempRootIntoExecutionBindings() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-cli-temp-binding-");
    Path requestPath = workspace.resolve("request.json");
    Path customTempRoot = workspace.resolve("custom-scratch");
    AtomicReference<Path> observedTempRoot = new AtomicReference<>();
    Files.writeString(
        requestPath, requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]"));

    int exitCode =
        GridGrindCli.forTesting(
                (request, bindings, sink) -> {
                  observedTempRoot.set(bindings.tempRoot());
                  return GridGrindResponses.success(List.of(), List.of(), List.of());
                })
            .run(
                new String[] {
                  "--request", requestPath.toString(), "--temp-root", customTempRoot.toString()
                },
                InputStream.nullInputStream(),
                new ByteArrayOutputStream());

    assertEquals(0, exitCode);
    assertEquals(customTempRoot.toAbsolutePath().normalize(), observedTempRoot.get());
  }

  @Test
  void verboseExecutionJournalStreamsLiveEventsToStderr() throws IOException {
    String request =
        requestJsonWithPlanId(
            "ledger-audit",
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            verboseExecutionJson(),
            emptyFormulaEnvironmentJson(),
            """
            [
              { "stepId": "ensure-ledger", "target": { "type": "SHEET_BY_NAME", "name": "Ledger" }, "action": { "type": "ENSURE_SHEET" } },
              { "stepId": "summary", "target": { "type": "WORKBOOK_CURRENT" }, "query": { "type": "GET_WORKBOOK_SUMMARY" } }
            ]
            """);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    GridGrindResponse.Success success =
        assertInstanceOf(
            GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(0, exitCode);
    assertEquals("ledger-audit", success.journal().planId().orElseThrow());
    assertTrue(
        stderr.toString(StandardCharsets.UTF_8).contains("[gridgrind]"),
        "verbose journal must emit live stderr lines");
    assertTrue(
        stderr.toString(StandardCharsets.UTF_8).contains("ensure-ledger"),
        "stderr must include the step id");
  }

  @Test
  void returnsStructuredJsonErrorForInvalidRequest() throws IOException {
    String request =
        requestJson(
            "{ \"type\": \"EXISTING\", \"path\": \"/tmp/does-not-exist.xlsx\" }",
            "{ \"type\": \"NONE\" }",
            "[]");

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    GridGrindResponse response = response(stdout, stderr);

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertTrue(failure.problem().message().contains("Workbook does not exist"));
  }

  @Test
  void reportsInvalidSheetCharactersDuringRequestRead() throws IOException {
    String request =
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              { "stepId": "ensure-bad-sheet", "target": { "type": "SHEET_BY_NAME", "name": "Bad:Name" }, "action": { "type": "ENSURE_SHEET" } }
            ]
            """);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.code());
    assertEquals("execute", failure.command());
    assertEquals(
        java.util.Optional.of("steps[0].target.name"), failure.location().orElseThrow().jsonPath());
    assertTrue(failure.message().contains("invalid Excel character ':'"));
  }

  @Test
  void acceptsExplicitSetTableRequestsWithShowTotalsRowFalse() throws IOException {
    String request =
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
                { "stepId": "ensure-dispatch", "target": { "type": "SHEET_BY_NAME", "name": "Dispatch" }, "action": { "type": "ENSURE_SHEET" } },
                {
                  "stepId": "seed-dispatch",
                  "target": { "type": "RANGE_BY_RANGE", "sheetName": "Dispatch", "range": "A1:B3" },
                  "action": {
                    "type": "SET_RANGE",
                    "rows": {
                      "type": "TYPED",
                      "cells": [
                        [
                          { "type": "TEXT", "source": { "type": "INLINE", "text": "Owner" } },
                          { "type": "TEXT", "source": { "type": "INLINE", "text": "Task" } }
                        ],
                        [
                          { "type": "TEXT", "source": { "type": "INLINE", "text": "Ada" } },
                          { "type": "TEXT", "source": { "type": "INLINE", "text": "Onboarding" } }
                        ],
                        [
                          { "type": "TEXT", "source": { "type": "INLINE", "text": "Lin" } },
                          { "type": "TEXT", "source": { "type": "INLINE", "text": "Badge run" } }
                        ]
                      ]
                    }
                  }
                },
                {
                  "stepId": "set-dispatch-table",
                  "target": { "type": "TABLE_BY_NAME_ON_SHEET", "name": "DispatchQueue", "sheetName": "Dispatch" },
                  "action": {
                    "type": "SET_TABLE",
                    "table": {
                      "name": "DispatchQueue",
                      "sheetName": "Dispatch",
                      "range": "A1:B3",
                      "showTotalsRow": false,
                      "hasAutofilter": true,
                      "style": { "type": "NONE" },
                      "comment": { "type": "INLINE", "text": "" },
                      "published": false,
                      "insertRow": false,
                      "insertRowShift": false,
                      "headerRowCellStyle": "",
                      "dataCellStyle": "",
                      "totalsRowCellStyle": "",
                      "columns": []
                    }
                  }
                },
                { "stepId": "tables", "target": { "type": "TABLE_ALL" }, "query": { "type": "GET_TABLES" } }
              ]
            """);

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments(),
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    WorkbookAssetInspectionResult.TablesResult tables =
        (WorkbookAssetInspectionResult.TablesResult) success.inspections().getFirst();
    assertEquals(1, tables.tables().size());
    assertEquals(0, tables.tables().getFirst().structure().totalsRowCount());
  }

  @Test
  void readsJsonRequestFromFileAndWritesJsonResponseToFile() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-request-", ".json");
    Path responsePath =
        Files.createTempDirectory("gridgrind-response-").resolve("nested").resolve("response.json");

    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              { "stepId": "ensure-budget", "target": { "type": "SHEET_BY_NAME", "name": "Budget" }, "action": { "type": "ENSURE_SHEET" } },
              { "stepId": "workbook", "target": { "type": "WORKBOOK_CURRENT" }, "query": { "type": "GET_WORKBOOK_SUMMARY" } }
            ]
            """));

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--request", requestPath.toString(), "--response", responsePath.toString()
                },
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);

    GridGrindResponse response = GridGrindJson.readResponse(Files.readAllBytes(responsePath));

    assertEquals(0, exitCode);
    assertEquals(0, stdout.size());
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals(
        List.of("Budget"),
        ((WorkbookInspectionResult.WorkbookSummaryResult) success.inspections().getFirst())
            .workbook()
            .sheetNames());
  }

  @Test
  void writesFailurePointerToStderrWhenExecutionFailsAndResponsePathIsConfigured()
      throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-invalid-request-", ".json");
    Path responsePath = Files.createTempFile("gridgrind-invalid-response-", ".json");
    Files.deleteIfExists(responsePath);
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              {
                "stepId": "summary",
                "target": { "type": "WORKBOOK" },
                "query": { "type": "GET_WORKBOOK_SUMMARY" }
              }
            ]
            """));

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--request", requestPath.toString(), "--response", responsePath.toString()
                },
                InputStream.nullInputStream(),
                stdout,
                stderr);

    CliFailureReport failure = cliFailure(Files.readAllBytes(responsePath));

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains(
                "GridGrind wrote the request failure report to " + responsePath.toAbsolutePath()));
    assertTrue(stderr.toString(StandardCharsets.UTF_8).contains("[INVALID_REQUEST_SHAPE:"));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.code());
    assertEquals(
        Optional.of("steps[0].target.type"),
        failure.location().flatMap(dev.erst.gridgrind.cli.discovery.CliFailureLocation::jsonPath));
  }
}
