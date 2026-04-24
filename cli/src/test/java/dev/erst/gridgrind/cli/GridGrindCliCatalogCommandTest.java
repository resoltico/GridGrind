package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GoalPlanReport;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindTaskCatalog;
import dev.erst.gridgrind.contract.catalog.GridGrindTaskPlanner;
import dev.erst.gridgrind.contract.catalog.TaskCatalog;
import dev.erst.gridgrind.contract.catalog.TaskPlanTemplate;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;
import org.junit.jupiter.api.Test;

/** Catalog and discovery command integration tests for GridGrindCli. */
class GridGrindCliCatalogCommandTest extends GridGrindCliTestSupport {
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
  void printExampleFlagSuggestsStableIdForCommonNonCanonicalTokens() throws IOException {
    assertSuggestedExampleId("chart");
    assertSuggestedExampleId("chart-request.json");
    assertSuggestedExampleId("chart-request");
    assertSuggestedExampleId("chart request");
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

  private static void assertSuggestedExampleId(String authoredValue) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example", authoredValue},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));
    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("--print-example", failure.problem().context().argument());
    assertTrue(
        failure.problem().message().contains("did you mean CHART?"),
        () -> "expected CHART suggestion for authored value " + authoredValue);
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
  void printTaskCatalogRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--version"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(2, exitCode);
    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("--version", failure.problem().context().argument());
    assertTrue(failure.problem().message().contains("Unknown argument"));
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
  void printTaskPlanRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-plan", "DASHBOARD", "--request", "ignored.json"},
                new ByteArrayInputStream(new byte[0]),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("--request", failure.problem().context().argument());
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
                          "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
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
                          "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                          "action": { "type": "ENSURE_SHEET" }
                        },
                        {
                          "stepId": "set-title",
                          "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
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
  void doctorRequestCanWriteReportToAnExplicitResponsePath() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-doctor-report-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--response", responsePath.toString()},
                new ByteArrayInputStream(
                    """
                    {
                      "source": { "type": "NEW" },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    RequestDoctorReport report =
        GridGrindJson.readRequestDoctorReport(Files.readAllBytes(responsePath));

    assertEquals(0, exitCode);
    assertTrue(report.valid());
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void doctorRequestResolvesRelativeInputsFromTheRequestFileDirectory() throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-doctor-root-");
    Path requestPath = requestDirectory.resolve("doctor request.json");
    Path payloadPath = requestDirectory.resolve("blank.txt");
    Files.writeString(payloadPath, "");
    Files.writeString(
        requestPath,
        """
        {
          "source": { "type": "NEW" },
          "steps": [
            {
              "stepId": "ensure-budget",
              "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
              "action": { "type": "ENSURE_SHEET" }
            },
            {
              "stepId": "set-title",
              "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
              "action": {
                "type": "SET_CELL",
                "value": {
                  "type": "TEXT",
                  "source": { "type": "UTF8_FILE", "path": "blank.txt" }
                }
              }
            }
          ]
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

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, report.problem().code());
    assertEquals("RESOLVE_INPUTS", report.problem().context().stage());
    assertEquals("cell text must not be blank", report.problem().message());
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
  void doctorRequestPreflightsExistingWorkbookSourcesFromTheRequestFileDirectory()
      throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-doctor-source-");
    Path requestPath = requestDirectory.resolve("doctor-existing.json");
    Files.writeString(
        requestPath,
        """
        {
          "source": { "type": "EXISTING", "path": "missing-workbook.xlsx" },
          "steps": [
            {
              "stepId": "summary",
              "target": { "type": "WORKBOOK_CURRENT" },
              "query": { "type": "GET_WORKBOOK_SUMMARY" }
            }
          ]
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

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, report.problem().code());
    assertEquals("OPEN_WORKBOOK", report.problem().context().stage());
    assertEquals(
        requestDirectory.resolve("missing-workbook.xlsx").toString(),
        report.problem().context().sourceWorkbookPath());
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
  void printProtocolCatalogCanWriteItsPayloadToAnExplicitResponsePath() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-protocol-catalog-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--response", responsePath.toString()},
                InputStream.nullInputStream(),
                stdout);

    Catalog catalog = GridGrindJson.readProtocolCatalog(Files.readAllBytes(responsePath));

    assertEquals(0, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(GridGrindProtocolCatalog.catalog(), catalog);
  }

  @Test
  void printProtocolCatalogWithUnexpectedTrailingArgReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--version"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(2, exitCode);
    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("--version"));
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
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.journal().outcome().failureCode());
    assertTrue(failure.problem().message().contains("BOGUS_XYZ"));
  }

  @Test
  void printProtocolCatalogWritesStructuredFailuresToTheResponsePathWhenConfigured()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-protocol-error-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-protocol-catalog",
                  "--operation",
                  "BOGUS_XYZ",
                  "--response",
                  responsePath.toString()
                },
                InputStream.nullInputStream(),
                stdout);

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            GridGrindJson.readResponse(Files.readAllBytes(responsePath)));

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("BOGUS_XYZ"));
  }

  @Test
  void printProtocolCatalogWithSearchReturnsRankedMatches() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-protocol-catalog", "--search", "sheet layout"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    String output = stdout.toString(StandardCharsets.UTF_8).trim();
    assertTrue(output.contains("\"query\" : \"sheet layout\""));
    assertTrue(output.contains("\"qualifiedId\" : \"inspectionQueryTypes:GET_SHEET_LAYOUT\""));
    assertTrue(output.contains("\"kind\" : \"ENTRY\""));
  }

  @Test
  void printProtocolCatalogRejectsOperationAndSearchTogether() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-protocol-catalog", "--operation", "SET_CELL", "--search", "cell"
                },
                InputStream.nullInputStream(),
                stdout);

    assertEquals(2, exitCode);
    GridGrindResponse response = GridGrindJson.readResponse(stdout.toByteArray());
    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("--operation"));
    assertTrue(failure.problem().message().contains("--search"));
  }
}
