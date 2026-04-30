package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.executor.ExecutionJournalSink;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
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

  @Test
  void threeArgumentConstructorStillRunsWithDefaultJournalWriter() throws IOException {
    String request = requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]");

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli(
                (ignoredRequest, ignoredBindings, ignoredSink) ->
                    GridGrindResponses.success(
                        null,
                        new GridGrindResponsePersistence.PersistenceOutcome.NotSaved(),
                        List.of(),
                        List.of(),
                        List.of()),
                new CliRequestReader(),
                new CliResponseWriter())
            .run(
                new String[0],
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    assertEquals(0, exitCode);
    assertInstanceOf(
        GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
  }

  @Test
  void fourArgumentConstructorStillUsesProvidedJournalWriter() throws IOException {
    String request = requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]");

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli(
                (ignoredRequest, ignoredBindings, ignoredSink) ->
                    GridGrindResponses.success(
                        null,
                        new GridGrindResponsePersistence.PersistenceOutcome.NotSaved(),
                        List.of(),
                        List.of(),
                        List.of()),
                new CliRequestReader(),
                new CliResponseWriter(),
                new CliJournalWriter())
            .run(
                new String[0],
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    assertEquals(0, exitCode);
    assertInstanceOf(
        GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
  }

  @Test
  void cliJournalWriterReturnsNoopWhenRequestIsMissing() {
    CliJournalWriter writer = new CliJournalWriter();

    assertSame(ExecutionJournalSink.NOOP, writer.sinkFor(null, OutputStream.nullOutputStream()));
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
              { "stepId": "append-header", "target": { "type": "SHEET_BY_NAME", "name": "Budget" }, "action": { "type": "APPEND_ROW", "values": [
                { "type": "TEXT", "source": { "type": "INLINE", "text": "Item" } },
                { "type": "TEXT", "source": { "type": "INLINE", "text": "Amount" } }
              ] } },
              { "stepId": "append-hosting", "target": { "type": "SHEET_BY_NAME", "name": "Budget" }, "action": { "type": "APPEND_ROW", "values": [
                { "type": "TEXT", "source": { "type": "INLINE", "text": "Hosting" } },
                { "type": "NUMBER", "number": 49.0 }
              ] } },
              { "stepId": "set-total", "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "B3" }, "action": { "type": "SET_CELL", "value": { "type": "FORMULA", "source": { "type": "INLINE", "text": "SUM(B2:B2)" } } } },
              { "stepId": "workbook", "target": { "type": "WORKBOOK_CURRENT" }, "query": { "type": "GET_WORKBOOK_SUMMARY" } },
              { "stepId": "cells", "target": { "type": "CELL_BY_ADDRESSES", "sheetName": "Budget", "addresses": ["A1", "B3"] }, "query": { "type": "GET_CELLS" } }
            ]
            """);

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
    assertEquals(List.of(), success.warnings());
    GridGrindWorkbookSurfaceReports.WorkbookSummary workbook =
        ((InspectionResult.WorkbookSummaryResult) success.inspections().get(0)).workbook();
    assertEquals("Budget", workbook.sheetNames().get(0));
    InspectionResult.CellsResult cells =
        (InspectionResult.CellsResult) success.inspections().get(1);
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport b3Cell =
        (dev.erst.gridgrind.contract.dto.CellReport.FormulaReport) cells.cells().get(1);
    assertEquals("SUM(B2:B2)", b3Cell.formula());
    assertEquals(
        49.0,
        ((dev.erst.gridgrind.contract.dto.CellReport.NumberReport) b3Cell.evaluation())
            .numberValue());
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
    int exitCode =
        new GridGrindCli()
            .run(
                new String[0],
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("PARSE_ARGUMENTS", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
    assertTrue(
        failure.problem().message().contains("STANDARD_INPUT-authored values require --request"));
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
    InspectionResult.CellsResult cells =
        assertInstanceOf(InspectionResult.CellsResult.class, success.inspections().getFirst());
    dev.erst.gridgrind.contract.dto.CellReport.TextReport a1 =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.TextReport.class, cells.cells().getFirst());
    assertEquals(0, exitCode);
    assertEquals("Quarterly Budget", a1.stringValue());
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
                new String[0],
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
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("READ_REQUEST", failure.problem().context().stage());
    assertEquals(java.util.Optional.of("steps[0].target"), readRequestContext(failure).jsonPath());
    assertTrue(failure.problem().message().contains("invalid Excel character ':'"));
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
                    "rows": [
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
                new String[0],
                new ByteArrayInputStream(request.getBytes(StandardCharsets.UTF_8)),
                stdout);

    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    InspectionResult.TablesResult tables =
        (InspectionResult.TablesResult) success.inspections().getFirst();
    assertEquals(1, tables.tables().size());
    assertEquals(0, tables.tables().getFirst().totalsRowCount());
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
        ((InspectionResult.WorkbookSummaryResult) success.inspections().getFirst())
            .workbook()
            .sheetNames());
  }
}
