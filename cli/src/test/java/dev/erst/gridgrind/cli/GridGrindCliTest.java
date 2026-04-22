package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.CliSurface;
import dev.erst.gridgrind.contract.catalog.GoalPlanReport;
import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindTaskCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindTaskPlanner;
import dev.erst.gridgrind.contract.catalog.ShippedExampleEntry;
import dev.erst.gridgrind.contract.catalog.TaskCatalog;
import dev.erst.gridgrind.contract.catalog.TaskPlanTemplate;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCategory;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.executor.ExecutionJournalSink;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.UUID;
import org.junit.jupiter.api.Test;

/** Integration tests for GridGrindCli command-line invocation. */
class GridGrindCliTest {
  @Test
  void threeArgumentConstructorStillRunsWithDefaultJournalWriter() throws IOException {
    String request =
        """
        {
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": []
        }
        """;

    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli(
                (ignoredRequest, ignoredBindings, ignoredSink) ->
                    new GridGrindResponse.Success(
                        null,
                        new GridGrindResponse.PersistenceOutcome.NotSaved(),
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
  void cliJournalWriterReturnsNoopWhenRequestIsMissing() {
    CliJournalWriter writer = new CliJournalWriter();

    assertSame(ExecutionJournalSink.NOOP, writer.sinkFor(null, OutputStream.nullOutputStream()));
  }

  @Test
  void cliJournalWriterSwallowsBestEffortIoFailures() throws IOException {
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
            {
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": {
                "journal": { "level": "VERBOSE" }
              },
              "steps": []
            }
            """
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
                          "2026-04-18T11:45:00Z", "OPEN", "opened", null, null)));
    }
  }

  @Test
  void readsJsonRequestFromStdinAndWritesJsonResponse() throws IOException {
    String request =
        """
            {
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": {
                "calculation": { "strategy": { "type": "EVALUATE_ALL" } }
              },
              "steps": [
                { "stepId": "ensure-budget", "target": { "type": "BY_NAME", "name": "Budget" }, "action": { "type": "ENSURE_SHEET" } },
                { "stepId": "append-header", "target": { "type": "BY_NAME", "name": "Budget" }, "action": { "type": "APPEND_ROW", "values": [
                  { "type": "TEXT", "source": { "type": "INLINE", "text": "Item" } },
                  { "type": "TEXT", "source": { "type": "INLINE", "text": "Amount" } }
                ] } },
                { "stepId": "append-hosting", "target": { "type": "BY_NAME", "name": "Budget" }, "action": { "type": "APPEND_ROW", "values": [
                  { "type": "TEXT", "source": { "type": "INLINE", "text": "Hosting" } },
                  { "type": "NUMBER", "number": 49.0 }
                ] } },
                { "stepId": "set-total", "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "B3" }, "action": { "type": "SET_CELL", "value": { "type": "FORMULA", "source": { "type": "INLINE", "text": "SUM(B2:B2)" } } } },
                { "stepId": "workbook", "target": { "type": "CURRENT" }, "query": { "type": "GET_WORKBOOK_SUMMARY" } },
                { "stepId": "cells", "target": { "type": "BY_ADDRESSES", "sheetName": "Budget", "addresses": ["A1", "B3"] }, "query": { "type": "GET_CELLS" } }
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
    assertEquals(List.of(), success.warnings());
    GridGrindResponse.WorkbookSummary workbook =
        ((InspectionResult.WorkbookSummaryResult) success.inspections().get(0)).workbook();
    assertEquals("Budget", workbook.sheetNames().get(0));
    InspectionResult.CellsResult cells =
        (InspectionResult.CellsResult) success.inspections().get(1);
    GridGrindResponse.CellReport.FormulaReport b3Cell =
        (GridGrindResponse.CellReport.FormulaReport) cells.cells().get(1);
    assertEquals("SUM(B2:B2)", b3Cell.formula());
    assertEquals(
        49.0, ((GridGrindResponse.CellReport.NumberReport) b3Cell.evaluation()).numberValue());
  }

  @Test
  void rejectsStandardInputAuthoredValuesWhenRequestAlsoUsesStdin() throws IOException {
    String request =
        """
            {
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [
                { "stepId": "ensure-budget", "target": { "type": "BY_NAME", "name": "Budget" }, "action": { "type": "ENSURE_SHEET" } },
                {
                  "stepId": "set-title",
                  "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                  "action": {
                    "type": "SET_CELL",
                    "value": {
                      "type": "TEXT",
                      "source": { "type": "STANDARD_INPUT" }
                    }
                  }
                }
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

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("PARSE_ARGUMENTS", failure.problem().context().stage());
    assertEquals("--request", failure.problem().context().argument());
    assertTrue(
        failure.problem().message().contains("STANDARD_INPUT-authored values require --request"));
  }

  @Test
  void bindsStandardInputToSourceBackedValuesWhenRequestComesFromFile() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-stdin-request-", ".json");
    Files.writeString(
        requestPath,
        """
        {
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            { "stepId": "ensure-budget", "target": { "type": "BY_NAME", "name": "Budget" }, "action": { "type": "ENSURE_SHEET" } },
            {
              "stepId": "set-title",
              "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
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
              "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
              "query": { "type": "GET_CELLS" }
            }
          ]
        }
        """,
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
    GridGrindResponse.CellReport.TextReport a1 =
        assertInstanceOf(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst());
    assertEquals(0, exitCode);
    assertEquals("Quarterly Budget", a1.stringValue());
  }

  @Test
  void verboseExecutionJournalStreamsLiveEventsToStderr() throws IOException {
    String request =
        """
            {
              "planId": "ledger-audit",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": {
                "journal": { "level": "VERBOSE" }
              },
              "steps": [
                { "stepId": "ensure-ledger", "target": { "type": "BY_NAME", "name": "Ledger" }, "action": { "type": "ENSURE_SHEET" } },
                { "stepId": "summary", "target": { "type": "CURRENT" }, "query": { "type": "GET_WORKBOOK_SUMMARY" } }
              ]
            }
            """;

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
    assertEquals("ledger-audit", success.journal().planId());
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
        """
            {
              "source": { "type": "EXISTING", "path": "/tmp/does-not-exist.xlsx" },
              "persistence": { "type": "NONE" },
              "steps": []
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
  void reportsInvalidSheetCharactersDuringRequestRead() throws IOException {
    String request =
        """
            {
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [
                { "stepId": "ensure-bad-sheet", "target": { "type": "BY_NAME", "name": "Bad:Name" }, "action": { "type": "ENSURE_SHEET" } }
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

    assertEquals(1, exitCode);
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("READ_REQUEST", failure.problem().context().stage());
    assertEquals("steps[0].target", failure.problem().context().jsonPath());
    assertTrue(failure.problem().message().contains("invalid Excel character ':'"));
  }

  @Test
  void acceptsSetTableRequestsThatOmitShowTotalsRow() throws IOException {
    String request =
        """
            {
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [
                { "stepId": "ensure-dispatch", "target": { "type": "BY_NAME", "name": "Dispatch" }, "action": { "type": "ENSURE_SHEET" } },
                {
                  "stepId": "seed-dispatch",
                  "target": { "type": "BY_RANGE", "sheetName": "Dispatch", "range": "A1:B3" },
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
                  "target": { "type": "BY_NAME_ON_SHEET", "name": "DispatchQueue", "sheetName": "Dispatch" },
                  "action": {
                    "type": "SET_TABLE",
                    "table": {
                      "name": "DispatchQueue",
                      "sheetName": "Dispatch",
                      "range": "A1:B3",
                      "style": { "type": "NONE" }
                    }
                  }
                },
                { "stepId": "tables", "target": { "type": "ALL" }, "query": { "type": "GET_TABLES" } }
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
        """
            {
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [
                { "stepId": "ensure-budget", "target": { "type": "BY_NAME", "name": "Budget" }, "action": { "type": "ENSURE_SHEET" } },
                { "stepId": "workbook", "target": { "type": "CURRENT" }, "query": { "type": "GET_WORKBOOK_SUMMARY" } }
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
        ((InspectionResult.WorkbookSummaryResult) success.inspections().getFirst())
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
    // The description comes from the processed gridgrind.properties resource on the test classpath.
    String output = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(output.startsWith("GridGrind unknown\n"), "must start with GridGrind unknown");
    assertTrue(output.endsWith("\n"), "must end with newline");
    assertTrue(output.lines().count() >= 2, "must have at least two lines");
  }

  @Test
  void licenseFlagPrintsLicenseTextToStdoutAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--license"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    String license = stdout.toString(StandardCharsets.UTF_8);
    // The license text is absent from the test classpath; expect the fallback notice.
    assertFalse(license.isBlank());
  }

  @Test
  void licenseText_containsMitLicense_whenResourcePresent() {
    String mit = "MIT License\n\nCopyright (c) 2026 Ervins Strauhmanis\n";
    InputStream own = new ByteArrayInputStream(mit.getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, null, null, null);

    assertTrue(result.contains("MIT License"));
    assertTrue(result.contains("Ervins Strauhmanis"));
    assertFalse(
        result.contains("Third-party notices and licenses:"), "no third-party section when absent");
  }

  @Test
  void licenseText_containsThirdPartySection_whenDependencyLicensesPresent() {
    InputStream own = new ByteArrayInputStream("MIT License\n".getBytes(StandardCharsets.UTF_8));
    InputStream notice = new ByteArrayInputStream("NOTICE info\n".getBytes(StandardCharsets.UTF_8));
    InputStream apache =
        new ByteArrayInputStream("Apache License\n".getBytes(StandardCharsets.UTF_8));
    InputStream bsd = new ByteArrayInputStream("BSD License\n".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, notice, apache, bsd);

    assertTrue(result.contains("MIT License"));
    assertTrue(result.contains("Third-party notices and licenses:"));
    assertTrue(result.contains("NOTICE info"));
    assertTrue(result.contains("Apache License"));
    assertTrue(result.contains("BSD License"));
  }

  @Test
  void licenseText_returnsFallback_whenAllResourcesAbsent() {
    String result = GridGrindCli.licenseText(null, null, null, null);

    assertFalse(result.isBlank());
    assertTrue(result.contains("not available"));
  }

  @Test
  void licenseText_thirdPartyOnly_whenOwnAbsent() {
    InputStream apache =
        new ByteArrayInputStream("Apache License\n".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(null, null, apache, null);

    assertTrue(result.contains("Apache License"));
    assertFalse(result.contains("---"), "no separator when own license is absent");
  }

  @Test
  void licenseText_ensuresTrailingNewline_whenContentLacksIt() {
    InputStream own = new ByteArrayInputStream("MIT License".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, null, null, null);

    assertTrue(result.endsWith("\n"), "must end with newline even when source text does not");
  }

  @Test
  void licenseText_skipsUnreadableStream() {
    // Pass the broken stream directly to avoid a PMD CloseResource warning;
    // append() closes it via try-with-resources even on IOException.
    String result =
        GridGrindCli.licenseText(
            new InputStream() {
              @Override
              public int read() throws IOException {
                throw new IOException("simulated read failure");
              }

              @Override
              public int read(byte[] buf, int off, int len) throws IOException {
                throw new IOException("simulated read failure");
              }
            },
            null,
            null,
            null);

    // The broken stream is skipped; all streams absent triggers the fallback.
    assertTrue(result.contains("not available"));
  }

  @Test
  void licenseText_containsNotice_whenNoticePresent() {
    InputStream own = new ByteArrayInputStream("MIT License\n".getBytes(StandardCharsets.UTF_8));
    InputStream notice =
        new ByteArrayInputStream("NOTICE content\n".getBytes(StandardCharsets.UTF_8));

    String result = GridGrindCli.licenseText(own, notice, null, null);

    assertTrue(result.contains("Third-party notices and licenses:"));
    assertTrue(result.contains("NOTICE content"));
  }

  @Test
  void productHeader_formatsVersionAndDescription() {
    assertEquals(
        "GridGrind 1.0.0\nA description", GridGrindCli.productHeader("1.0.0", "A description"));
  }

  @Test
  void helpFlagsPrintUsageAndReturnExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int longExitCode =
        new GridGrindCli()
            .run(new String[] {"--help"}, new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, longExitCode);
    String help = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(help.contains("GridGrind"));
    assertTrue(help.contains("Usage:"));
    assertTrue(help.contains("Minimal Valid Request:"));
    assertTrue(help.contains("--request <path>"));
    assertTrue(help.contains("--print-request-template"));
    assertTrue(help.contains("--print-task-catalog"));
    assertTrue(help.contains("--print-task-plan <id>"));
    assertTrue(help.contains("--print-goal-plan <goal>"));
    assertTrue(help.contains("--doctor-request"));
    assertTrue(help.contains("--print-protocol-catalog"));
    assertTrue(help.contains("--print-example <id>"));
    assertTrue(help.contains("--help, -h"));
    assertTrue(help.contains("blob/main/docs/QUICK_REFERENCE.md"));
    assertTrue(help.contains("Coordinate Systems:"));
    assertTrue(
        help.contains(
            GridGrindProtocolCatalog.catalog()
                .cliSurface()
                .fileWorkflow()
                .entries()
                .getLast()
                .value()));
    assertTrue(
        help.contains("type group accepted by polymorphic fields"),
        "help must explain what the protocol catalog publishes");
    assertTrue(
        help.contains("high-level office-work recipes"),
        "help must explain what the task catalog publishes");
    assertTrue(
        help.contains("cellInputTypes:FORMULA"),
        "help must explain how to qualify duplicate catalog ids");

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

    assertTrue(help.contains("GridGrind 0.9.0"));
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
    assertTrue(help.contains("GridGrind 1.0.0"), "help must contain the version line");
    // The description line appears immediately after the version line.
    int versionLineEnd = help.indexOf("GridGrind 1.0.0") + "GridGrind 1.0.0".length();
    String afterVersion = help.substring(versionLineEnd).stripLeading();
    assertFalse(
        afterVersion.startsWith("Usage:"),
        "A description line must appear between the version and Usage:");
  }

  @Test
  void helpTextDockerExampleUsesMountedWorkingDirectoryPaths() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(help.contains("-w /workdir"), "Docker example must show -w /workdir");
    assertTrue(
        help.contains("--request request.json"),
        "Docker example must use request paths relative to the mounted workdir");
    assertTrue(
        help.contains("--response response.json"),
        "Docker example must use response paths relative to the mounted workdir");
  }

  @Test
  void helpTextExplainsFileWorkflow() {
    String help = GridGrindCli.helpText("1.0.0");
    CliSurface cliSurface = GridGrindProtocolCatalog.catalog().cliSurface();

    assertTrue(help.contains("File Workflow:"));
    for (CliSurface.DefinitionEntry entry : cliSurface.fileWorkflow().entries()) {
      assertTrue(
          help.contains(entry.label() + ":"),
          () -> "help must include file workflow label: " + entry.label());
      assertTrue(
          help.contains(entry.value()),
          () -> "help must include file workflow value: " + entry.value());
    }
  }

  @Test
  void helpTextIncludesCoordinateSystemsTable() {
    String help = GridGrindCli.helpText("1.0.0");
    CliSurface cliSurface = GridGrindProtocolCatalog.catalog().cliSurface();

    assertTrue(help.contains("Coordinate Systems:"));
    for (CliSurface.CoordinateSystemEntry entry : cliSurface.coordinateSystems().entries()) {
      assertTrue(
          help.contains(entry.pattern()),
          () -> "help must include coordinate pattern " + entry.pattern());
      assertTrue(
          help.contains(entry.convention()),
          () -> "help must include coordinate convention " + entry.convention());
    }
  }

  @Test
  void helpTextListsBuiltInGeneratedExamples() {
    String help = GridGrindCli.helpText("1.0.0");
    CliSurface cliSurface = GridGrindProtocolCatalog.catalog().cliSurface();

    assertTrue(help.contains("Built-in generated examples:"));
    assertTrue(help.contains(cliSurface.discovery().printOneExampleCommand()));
    assertTrue(help.contains(GridGrindContractText.workbookFindingsDiscoverySummary()));
    assertTrue(help.contains(GridGrindContractText.stepKindSummary()));
    for (ShippedExampleEntry example : GridGrindProtocolCatalog.catalog().shippedExamples()) {
      assertTrue(help.contains(example.id()), () -> "help must include example id " + example.id());
      assertTrue(
          help.contains("examples/" + example.fileName()),
          () -> "help must include example file " + example.fileName());
    }
  }

  @Test
  void helpTextDocumentsFormulaAuthoringBoundaries() {
    String help = GridGrindCli.helpText("1.0.0");

    assertTrue(help.contains("Formula authoring:"), "help must include formula-authoring label");
    assertTrue(help.contains(GridGrindContractText.formulaAuthoringLimitSummary()));
    assertTrue(
        help.contains("Loaded formula support:"), "help must include loaded-formula-support label");
    assertTrue(help.contains(GridGrindContractText.loadedFormulaSupportSummary()));
  }

  @Test
  void helpTextIncludesStructuralEditLimitNotes() {
    String help = GridGrindCli.helpText("1.0.0");
    for (CliSurface.DefinitionEntry entry :
        GridGrindProtocolCatalog.catalog().cliSurface().limits().entries()) {
      if ("Row structural edits".equals(entry.label())
          || "Column structural edits".equals(entry.label())
          || "Chart mutations".equals(entry.label())
          || "Chart title formulas".equals(entry.label())
          || "Drawing validation".equals(entry.label())) {
        assertTrue(
            help.contains(entry.label() + ":"),
            () -> "help must include limit label: " + entry.label());
        assertTrue(
            help.contains(entry.value()), () -> "help must include limit value: " + entry.value());
      }
    }
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

          @Override
          public int read(byte[] bytes, int offset, int length) throws IOException {
            throw new IOException("simulated read failure");
          }
        }) {
      assertEquals("GridGrind", GridGrindCli.descriptionFrom(broken));
    }
  }

  @Test
  void requestTemplateTextRendersUtf8TemplateBytes() {
    assertEquals(
        "{\"protocolVersion\":\"V1\"}",
        GridGrindCli.requestTemplateText(
            () -> "{\"protocolVersion\":\"V1\"}".getBytes(StandardCharsets.UTF_8)));
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

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindProtocolCatalog.requestTemplate(), request);
  }

  @Test
  void printExampleFlagPrintsKnownGeneratedExampleAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example", "WORKBOOK_HEALTH"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(
        GridGrindProtocolCatalog.exampleFor("WORKBOOK_HEALTH").orElseThrow().plan(), request);
  }

  @Test
  void printExampleFlagRejectsUnknownExampleId() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example", "BOGUS_EXAMPLE"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));
    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("--print-example", failure.problem().context().argument());
    assertTrue(failure.problem().message().contains("BOGUS_EXAMPLE"));
  }

  @Test
  void printTaskCatalogFlagPrintsCurrentCatalogAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    TaskCatalog catalog = GridGrindJson.readTaskCatalog(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindTaskCatalog.catalog(), catalog);
  }

  @Test
  void printTaskCatalogWithTaskFilterReturnsMatchingEntry() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--task", "DASHBOARD"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"DASHBOARD\""), "output must contain the task id");
    assertTrue(output.contains("\"capabilityRefs\""), "output must contain phased capability refs");
    assertTrue(
        output.contains("\"mutationActionTypes\""),
        "output must expose the exact referenced protocol groups");
  }

  @Test
  void printTaskCatalogWithUnknownTaskReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--task", "BOGUS_TASK"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(2, exitCode);
    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("--task", failure.problem().context().argument());
    assertTrue(failure.problem().message().contains("BOGUS_TASK"));
  }

  @Test
  void printTaskCatalogWithNonTaskTrailingArgReturnsFullCatalog() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--version"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"tasks\""), "full task catalog must contain the tasks key");
    assertTrue(
        output.contains("\"DASHBOARD\""), "full task catalog must contain the dashboard task");
  }

  @Test
  void printTaskPlanFlagPrintsStarterTemplateAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-plan", "DASHBOARD"},
                InputStream.nullInputStream(),
                stdout);

    TaskPlanTemplate template = GridGrindJson.readTaskPlanTemplate(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindTaskPlanner.templateFor("DASHBOARD"), template);
  }

  @Test
  void printTaskPlanWithUnknownTaskReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-plan", "BOGUS_TASK"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(2, exitCode);
    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("--print-task-plan", failure.problem().context().argument());
    assertTrue(failure.problem().message().contains("BOGUS_TASK"));
  }

  @Test
  void printTaskPlanFlagTakesPrecedenceOverExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-plan", "DASHBOARD", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    TaskPlanTemplate template = GridGrindJson.readTaskPlanTemplate(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals("DASHBOARD", template.task().id());
  }

  @Test
  void printGoalPlanFlagPrintsRankedTaskMatchesAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-goal-plan", "monthly sales dashboard with charts"},
                InputStream.nullInputStream(),
                stdout);

    GoalPlanReport report = GridGrindJson.readGoalPlanReport(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals("monthly sales dashboard with charts", report.goal());
    assertEquals("DASHBOARD", report.candidates().getFirst().task().id());
    assertTrue(report.candidates().getFirst().matchedTerms().contains("dashboard"));
    assertTrue(report.candidates().getFirst().matchedTerms().contains("chart"));
  }

  @Test
  void printGoalPlanRejectsBlankGoalStrings() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--print-goal-plan", ""}, InputStream.nullInputStream(), stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("--print-goal-plan", failure.problem().context().argument());
  }

  @Test
  void doctorRequestFlagPrintsStructuredReportWithoutExecutingTheRequest() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli(
                (ignoredRequest, ignoredBindings, ignoredSink) -> {
                  throw new AssertionError("doctoring a request must not execute it");
                })
            .run(
                new String[] {"--doctor-request"},
                new ByteArrayInputStream(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "ensure-budget",
                          "target": { "type": "BY_NAME", "name": "Budget" },
                          "action": { "type": "ENSURE_SHEET" }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    RequestDoctorReport report = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertTrue(report.valid());
    assertEquals("NEW", report.summary().sourceType());
    assertEquals(1, report.summary().stepCount());
    assertEquals(1, report.summary().mutationStepCount());
  }

  @Test
  void doctorRequestReturnsStructuredInvalidJsonReport() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request"},
                new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8)),
                stdout);

    RequestDoctorReport report = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(GridGrindProblemCode.INVALID_JSON, report.problem().code());
    assertNull(report.summary());
  }

  @Test
  void doctorRequestRejectsImpossibleStandardInputBindingWhenRequestAlsoUsesStdin()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request"},
                new ByteArrayInputStream(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": [
                        {
                          "stepId": "ensure-budget",
                          "target": { "type": "BY_NAME", "name": "Budget" },
                          "action": { "type": "ENSURE_SHEET" }
                        },
                        {
                          "stepId": "set-title",
                          "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                          "action": {
                            "type": "SET_CELL",
                            "value": { "type": "TEXT", "source": { "type": "STANDARD_INPUT" } }
                          }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    RequestDoctorReport report = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertTrue(report.summary().requiresStandardInputBinding());
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, report.problem().code());
    assertEquals("--request", report.problem().context().argument());
  }

  @Test
  void doctorRequestCanReadTheRequestFromAFile() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-doctor-request-", ".json");
    Files.writeString(
        requestPath,
        """
        {
          "source": { "type": "NEW" },
          "steps": []
        }
        """);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout);

    RequestDoctorReport report = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertTrue(report.valid());
    assertFalse(report.summary().requiresStandardInputBinding());
  }

  @Test
  void doctorRequestReturnsStructuredInvalidReportForSemanticallyInvalidRequests()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request"},
                new ByteArrayInputStream(
                    """
                    {
                      "source": { "type": "NEW" },
                      "persistence": { "type": "OVERWRITE" },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    RequestDoctorReport report = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, report.problem().code());
    assertEquals("NEW", report.summary().sourceType());
  }

  @Test
  void doctorRequestReturnsStructuredReadErrorsWhenTheRequestFileCannotBeOpened()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    Path missingRequestPath =
        Path.of("tmp", "doctor-missing-" + UUID.randomUUID() + ".json").toAbsolutePath();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", missingRequestPath.toString()},
                InputStream.nullInputStream(),
                stdout);

    RequestDoctorReport report = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(GridGrindProblemCode.IO_ERROR, report.problem().code());
    assertNull(report.summary());
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

    Catalog catalog = GridGrindJson.readProtocolCatalog(stdout.toByteArray());

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
    String output = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(output.startsWith("GridGrind unknown\n"), "must start with GridGrind unknown");
    assertTrue(output.endsWith("\n"), "must end with newline");
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
    assertTrue(stdout.toString(StandardCharsets.UTF_8).contains("GridGrind"));
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

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindProtocolCatalog.requestTemplate(), request);
  }

  @Test
  void printExampleFlagTakesPrecedenceOverExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example", "ASSERTION", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindProtocolCatalog.exampleFor("ASSERTION").orElseThrow().plan(), request);
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
    assertEquals("NEW", failure.problem().context().sourceType());
    assertEquals("NONE", failure.problem().context().persistenceType());
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
    assertEquals("EXISTING", failure.problem().context().sourceType());
    assertEquals("OVERWRITE", failure.problem().context().persistenceType());
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
    assertEquals("EXISTING", failure.problem().context().sourceType());
    assertEquals("SAVE_AS", failure.problem().context().persistenceType());
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
                  "steps": [
                    { "stepId": "summary", "target": { "type": "CURRENT" } }
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
    assertEquals("steps[0]", failure.problem().context().jsonPath());
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
                        "type": "RECTANGULAR_WINDOW",
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
    assertEquals("steps[0].target", failure.problem().context().jsonPath());
    assertEquals(14, failure.problem().context().jsonLine());
    assertNotNull(failure.problem().context().jsonColumn());
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
        responseDirectory.toAbsolutePath().toString(), failure.problem().context().responsePath());
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
                    { "stepId": "ensure-data", "target": { "type": "BY_NAME", "name": "Data" }, "action": { "type": "ENSURE_SHEET" } },
                    { "stepId": "set-formula", "target": { "type": "BY_ADDRESS", "sheetName": "Data", "address": "A1" }, "action": { "type": "SET_CELL", "value": { "type": "FORMULA", "source": { "type": "INLINE", "text": "SUM(" } } } }
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
              "steps": []
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
                      "steps": [
                        { "stepId": "ensure-sheet", "target": { "type": "BY_NAME", "name": "S" }, "action": { "type": "ENSURE_SHEET" } },
                        {
                          "stepId": "w",
                          "target": {
                            "type": "RECTANGULAR_WINDOW",
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
    assertTrue(help.contains("409"), "help must state the row height limit");
    assertTrue(
        help.contains("AREA, AREA_3D, BAR, BAR_3D")
            && help.contains("SURFACE_3D")
            && help.contains("DOUGHNUT"),
        "help must state the authored chart boundary");
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
    assertTrue(
        output.contains("mutationActionTypes"),
        "full catalog must contain mutationActionTypes key");
    assertTrue(
        output.contains("inspectionQueryTypes"),
        "full catalog must contain inspectionQueryTypes key");
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
    assertTrue(
        output.contains("\"fields\""), "filtered catalog output must contain field descriptors");
    assertTrue(
        output.contains("\"targetSelectors\""),
        "filtered catalog output must expose allowed target selector families");
    assertTrue(
        output.contains("\"CellSelector\""),
        "filtered catalog output must identify the target selector family");
  }

  @Test
  void printProtocolCatalogWithQualifiedOperationFilterReturnsMatchingNestedEntry()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--operation", "cellInputTypes:FORMULA"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"FORMULA\""), "output must contain the entry id");
    assertTrue(output.contains("\"source\""), "qualified lookup must expose the source field");
    assertFalse(
        output.contains("\"refersToFormula\""),
        "qualified lookup must not silently return the named-range report variant");
  }

  @Test
  void printProtocolCatalogWithNestedGroupFilterReturnsMatchingNestedGroup() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-protocol-catalog", "--operation", "nestedTypes:cellInputTypes"
                },
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"group\" : \"cellInputTypes\""));
    assertTrue(output.contains("\"discriminatorField\" : \"type\""));
    assertTrue(output.contains("\"TEXT\""));
  }

  @Test
  void printProtocolCatalogWithPlainGroupFilterReturnsMatchingPlainGroup() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--operation", "chartInputType"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"group\" : \"chartInputType\""));
    assertTrue(output.contains("\"ChartInput\""));
    assertTrue(output.contains("\"plots\""));
  }

  @Test
  void printProtocolCatalogWithAmbiguousOperationReturnsErrorAndCandidates() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--operation", "FORMULA"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(2, exitCode);
    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("Ambiguous operation: FORMULA"));
    assertTrue(failure.problem().message().contains("cellInputTypes:FORMULA"));
    assertTrue(failure.problem().message().contains("namedRangeReportTypes:FORMULA"));
  }

  @Test
  void printProtocolCatalogWithSheetLayoutFilterMentionsPresentation() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--operation", "GET_SHEET_LAYOUT"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"GET_SHEET_LAYOUT\""), "output must contain the entry id");
    assertTrue(output.contains("presentation"), "summary must mention layout.presentation");
    assertTrue(output.contains("outlineLevel"), "summary must mention row/column outline state");
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
        help.contains("execution is optional"), "help must mention that execution is optional");
    assertTrue(
        help.contains("formulaEnvironment is optional"),
        "help must mention that formulaEnvironment is optional");
    assertTrue(help.contains("steps is optional"), "help must mention that steps is optional");
    assertTrue(
        help.contains("ASSERTION steps for first-class verification"),
        "help must mention assertion steps");
    assertTrue(
        help.contains("do not send step.type"),
        "help must explain that step kind is inferred instead of authored as step.type");
    assertTrue(
        help.contains("mutations, assertions, and inspections may be interleaved"),
        "help must describe the ordered step model");
    assertTrue(help.contains("EVENT_READ mode"), "help must describe EVENT_READ mode limits");
    assertTrue(
        help.contains("STREAMING_WRITE mode"), "help must describe STREAMING_WRITE mode limits");
    assertTrue(
        help.contains(GridGrindContractText.eventReadInspectionQueryTypePhrase()),
        "help must describe the canonical EVENT_READ read surface");
    assertTrue(
        help.contains(GridGrindContractText.streamingWriteMutationActionTypePhrase()),
        "help must expose the canonical streaming-write operation name");
    assertFalse(
        help.contains("FORCE_FORMULA_RECALC_ON_OPEN"),
        "help must not expose the rejected legacy shorthand");
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
