package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.ShippedExampleCatalog;
import dev.erst.gridgrind.cli.discovery.TaskCatalog;
import dev.erst.gridgrind.cli.discovery.TaskKeywordMatchReport;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
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
import java.util.List;
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
                new String[] {"--print-example", "--lookup", "WORKBOOK_HEALTH"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindShippedExamples.find("WORKBOOK_HEALTH").orElseThrow().plan(), request);
    assertFalse(stdout.toString(StandardCharsets.UTF_8).contains(": null"));
  }

  @Test
  void printExampleCatalogFlagPrintsCurrentExampleCatalogAndReturnsExitCodeZero()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example-catalog"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    ShippedExampleCatalog catalog =
        GridGrindCliJson.readShippedExampleCatalog(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindShippedExamples.catalog(), catalog);
    assertEquals(
        dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode.SELF_CONTAINED,
        catalog.examples().stream()
            .filter(example -> "BUDGET".equals(example.id()))
            .findFirst()
            .orElseThrow()
            .workspaceMode());
    assertEquals(
        dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
        catalog.examples().stream()
            .filter(example -> "PACKAGE_SECURITY_INSPECTION".equals(example.id()))
            .findFirst()
            .orElseThrow()
            .workspaceMode());
    assertEquals(
        java.util.List.of("package-security-assets/gridgrind-package-security.xlsx"),
        catalog.examples().stream()
            .filter(example -> "PACKAGE_SECURITY_INSPECTION".equals(example.id()))
            .findFirst()
            .orElseThrow()
            .requiredPaths());
  }

  @Test
  void printAssetBackedExampleWarnsOnStderrBeforeExecution() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example", "--lookup", "PACKAGE_SECURITY_INSPECTION"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(
        GridGrindShippedExamples.find("PACKAGE_SECURITY_INSPECTION").orElseThrow().plan(), request);
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains("requires copied asset paths beside the request file"),
        "asset-backed example printing must warn about copied assets");
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains("package-security-assets/gridgrind-package-security.xlsx"),
        "asset-backed example printing must name the required copied asset paths");
  }

  @Test
  void printExampleFlagRejectsUnknownExampleId() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example", "--lookup", "BOGUS_EXAMPLE"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--lookup"), failure.argument());
    assertTrue(failure.message().contains("BOGUS_EXAMPLE"));
    assertTrue(failure.message().contains("--print-example-catalog"));
  }

  @Test
  void printExampleFlagWritesUnknownExampleFailureToResponsePathAndPointsStderrAtIt()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-unknown-example-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-example",
                  "--lookup",
                  "BOGUS_EXAMPLE",
                  "--response",
                  responsePath.toString()
                },
                InputStream.nullInputStream(),
                stdout,
                stderr);

    CliFailureReport failure = cliFailure(Files.readAllBytes(responsePath));

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains(
                "GridGrind wrote the CLI failure report to " + responsePath.toAbsolutePath()));
    assertTrue(stderr.toString(StandardCharsets.UTF_8).contains("[INVALID_ARGUMENTS:"));
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertTrue(failure.message().contains("BOGUS_EXAMPLE"));
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

    TaskCatalog catalog = GridGrindCliJson.readTaskCatalog(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindTaskCatalog.catalog(), catalog);
  }

  private static void assertSuggestedExampleId(String authoredValue) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-example", "--lookup", authoredValue},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stdout); // stdout passed for both streams; use the bytes overload directly
    CliFailureReport failure = cliFailure(stdout.toByteArray());
    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--lookup"), failure.argument());
    assertTrue(
        failure.message().contains("did you mean CHART?"),
        () -> "expected CHART suggestion for authored value " + authoredValue);
    assertTrue(
        failure.message().contains("--print-example-catalog"),
        () -> "expected recovery guidance for authored value " + authoredValue);
  }

  @Test
  void printTaskCatalogWithTaskFilterReturnsMatchingEntry() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--lookup", "DASHBOARD"},
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
  void printTaskCatalogRejectsBlankTaskFilterWithStructuredFailure() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--lookup", ""},
                InputStream.nullInputStream(),
                stdout,
                stderr);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals("print-task-catalog", failure.command());
    assertEquals(java.util.Optional.of("--lookup"), failure.argument());
    assertTrue(failure.message().contains("task lookup id must not be blank"));
  }

  @Test
  void printTaskCatalogWithUnknownTaskReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--lookup", "BOGUS_TASK"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--lookup"), failure.argument());
    assertTrue(failure.message().contains("BOGUS_TASK"));
    assertTrue(failure.message().contains("--print-task-catalog"));
    assertTrue(
        failure
            .message()
            .contains("--print-task-keyword-match --query \"monthly sales dashboard\""));
  }

  @Test
  void printTaskCatalogWithNonCanonicalTaskIdSuggestsStableToken() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--lookup", "dashboard"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--lookup"), failure.argument());
    assertTrue(failure.message().contains("did you mean DASHBOARD?"));
    assertTrue(failure.message().contains("--print-task-catalog"));
    assertTrue(
        failure
            .message()
            .contains("--print-task-keyword-match --query \"monthly sales dashboard\""));
  }

  @Test
  void printTaskCatalogWithNormalizedTaskIdSuggestsStableToken() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--lookup", "audit existing workbook"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--lookup"), failure.argument());
    assertTrue(failure.message().contains("did you mean AUDIT_EXISTING_WORKBOOK?"));
    assertTrue(failure.message().contains("--print-task-catalog"));
    assertTrue(
        failure
            .message()
            .contains("--print-task-keyword-match --query \"monthly sales dashboard\""));
  }

  @Test
  void printTaskCatalogRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-catalog", "--version"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--version"), failure.argument());
    assertTrue(
        failure
            .message()
            .contains(
                "Only one primary command may be used per invocation; --print-task-catalog"
                    + " cannot be combined with --version"));
  }

  @Test
  void printTaskPlanFlagPrintsRunnableStarterRequestAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-plan", "--lookup", "DASHBOARD"},
                InputStream.nullInputStream(),
                stdout);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindTaskPlanner.requestFor("DASHBOARD"), request);
  }

  @Test
  void printAssetBackedTaskStarterWarnsOnStderrBeforeExecution() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-plan", "--lookup", "AUDIT_EXISTING_WORKBOOK"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindTaskPlanner.requestFor("AUDIT_EXISTING_WORKBOOK"), request);
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains("requires copied asset paths beside the request file"),
        "asset-backed task starter printing must warn about copied assets");
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains("task-starter-assets/workbook-ops-source.xlsx"),
        "asset-backed task starter printing must name the required copied asset paths");
  }

  @Test
  void printTaskPlanWithUnknownTaskReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-plan", "--lookup", "BOGUS_TASK"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--lookup"), failure.argument());
    assertTrue(failure.message().contains("BOGUS_TASK"));
    assertTrue(failure.message().contains("--print-task-catalog"));
    assertTrue(
        failure
            .message()
            .contains("--print-task-keyword-match --query \"monthly sales dashboard\""));
  }

  @Test
  void printTaskPlanRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-task-plan", "--lookup", "DASHBOARD", "--request", "ignored.json"
                },
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--request"), failure.argument());
  }

  @Test
  void printTaskKeywordMatchFlagPrintsRankedTaskMatchesAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-task-keyword-match", "--query", "monthly sales dashboard with charts"
                },
                InputStream.nullInputStream(),
                stdout);

    TaskKeywordMatchReport report =
        GridGrindCliJson.readTaskKeywordMatchReport(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals("monthly sales dashboard with charts", report.query());
    assertEquals("DASHBOARD", report.candidates().getFirst().taskId());
    assertTrue(report.candidates().getFirst().matchedTerms().contains("dashboard"));
    assertTrue(report.candidates().getFirst().matchedTerms().contains("chart"));
  }

  @Test
  void printTaskKeywordMatchReturnsNoCandidatesForGibberishQueries() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-keyword-match", "--query", "zzzz no such workflow"},
                InputStream.nullInputStream(),
                stdout);

    TaskKeywordMatchReport report =
        GridGrindCliJson.readTaskKeywordMatchReport(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(List.of("zzzz"), report.normalizedTerms());
    assertEquals(List.of("zzzz"), report.unmatchedTerms());
    assertEquals(List.of(), report.suggestedIntentTags());
    assertEquals(List.of(), report.candidates());
  }

  @Test
  void printTaskKeywordMatchRejectsBlankQueries() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-keyword-match", "--query", ""},
                InputStream.nullInputStream(),
                stdout,
                stderr);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--query"), failure.argument());
  }

  @Test
  void printTaskKeywordMatchRejectsQueriesThatNormalizeToNoSearchableTerms() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-task-keyword-match", "--query", "a"},
                InputStream.nullInputStream(),
                stdout,
                stderr);
    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals(java.util.Optional.of("--query"), failure.argument());
    assertTrue(failure.message().contains("searchable term after normalization"));
  }

  @Test
  void printRequestTemplatePrintsTheCurrentMachineReadableTemplate() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--print-request-template"}, InputStream.nullInputStream(), stdout);

    dev.erst.gridgrind.contract.dto.WorkbookPlan template =
        GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals("FULL_XSSF", template.execution().mode().modeType());
  }

  @Test
  void doctorRequestFlagPrintsStructuredReportWithoutExecutingTheRequest() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-stdin-success-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        GridGrindCli.forTesting(
                (ignoredRequest, ignoredBindings, ignoredSink) -> {
                  throw new AssertionError("doctoring a request must not execute it");
                })
            .run(
                new String[] {"--doctor-request", "--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            """
                            [
                              {
                                "stepId": "ensure-budget",
                                "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                                "action": { "type": "ENSURE_SHEET" }
                              }
                            ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    RequestDoctorReport report = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertTrue(report.valid());
    assertEquals("NEW", report.summary().orElseThrow().sourceType());
    assertEquals(1, report.summary().orElseThrow().stepCount());
    assertEquals(1, report.summary().orElseThrow().mutationStepCount());
  }

  @Test
  void doctorRequestReturnsCompactReadFailureForInvalidJson() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-invalid-json-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--execution-root", workspace.toString()},
                new ByteArrayInputStream("{".getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(GridGrindProblemCode.INVALID_JSON, report.primaryProblem().orElseThrow().code());
    assertEquals(java.util.Optional.of(1), readRequestContext(report).jsonLine());
    assertEquals(java.util.Optional.of(2), readRequestContext(report).jsonColumn());
  }

  @Test
  void doctorRequestRejectsMissingRequestInputWithCompactCliFailure() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--doctor-request"}, InputStream.nullInputStream(), stdout, stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals("doctor-request", failure.command());
    assertEquals(java.util.Optional.of("--request"), failure.argument());
    assertTrue(failure.message().contains("No request JSON was provided."));
  }

  @Test
  void doctorRequestRejectsImpossibleStandardInputBindingWhenRequestAlsoUsesStdin()
      throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-stdin-binding-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            """
                            [
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
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, report.primaryProblem().orElseThrow().code());
    assertTrue(report.primaryProblem().orElseThrow().message().contains("STANDARD_INPUT"));
  }

  @Test
  void doctorRequestBatchesIndependentProblemsWhenOneRequestHasMultipleDefects()
      throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-batch-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "EXISTING" },
                      "persistence": { "type": "SAVE_AS" },
                      "execution": {
                        "mode": { "type": "EVENT_READ" },
                        "journal": { "level": "NORMAL" },
                        "calculation": {
                          "strategy": { "type": "EVALUATE_ALL" },
                          "markRecalculateOnOpen": true
                        }
                      },
                      "formulaEnvironment": {
                        "externalWorkbooks": [],
                        "missingWorkbookPolicy": "ERROR",
                        "udfToolpacks": []
                      },
                      "steps": [
                        {
                          "stepId": "duplicate-step",
                          "target": { "type": "WORKBOOK_CURRENT" },
                          "query": { "type": "GET_WORKBOOK_SUMMARY" }
                        },
                        {
                          "stepId": "duplicate-step",
                          "target": {
                            "type": "CELL_BY_ADDRESS",
                            "sheetName": "Summary",
                            "address": "A1"
                          },
                          "assertion": {
                            "type": "EXPECT_CELL_VALUE",
                            "expectedValue": { "type": "TEXT", "text": "x" }
                          }
                        }
                      ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);
    List<String> problemMessages =
        report.problems().stream().map(GridGrindProblemDetail.Problem::message).toList();

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(3, report.problems().size());
    assertTrue(problemMessages.contains("Missing required field 'source.path'"));
    assertTrue(problemMessages.contains("Missing required field 'persistence.path'"));
    assertTrue(
        problemMessages.stream()
            .anyMatch(message -> message.contains("duplicate stepId values: duplicate-step")));
  }

  @Test
  void doctorRequestCanReadTheRequestFromAFile() throws IOException {
    Path requestPath = Files.createTempFile("gridgrind-doctor-request-", ".json");
    Files.writeString(
        requestPath, requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]"));
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
    assertFalse(report.summary().orElseThrow().requiresStandardInputBinding());
  }

  @Test
  void doctorRequestCanWriteReportToAnExplicitResponsePath() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-report-");
    Path responsePath = Files.createTempFile("gridgrind-doctor-report-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--doctor-request",
                  "--execution-root",
                  workspace.toString(),
                  "--response",
                  responsePath.toString()
                },
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    RequestDoctorReport report =
        GridGrindJson.readRequestDoctorReport(Files.readAllBytes(responsePath));

    assertEquals(0, exitCode);
    assertTrue(report.valid());
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void doctorRequestResolvesRelativeInputsFromTheRequestFileDirectory() throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-doctor-root-");
    Path requestPath = requestDirectory.resolve("doctor request.json");
    Path payloadPath = requestDirectory.resolve("blank.txt");
    Files.writeString(payloadPath, "");
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"NEW\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
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
            """));
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, report.primaryProblem().orElseThrow().code());
    assertEquals("RESOLVE_INPUTS", report.primaryProblem().orElseThrow().context().stage());
    assertEquals("cell text must not be blank", report.primaryProblem().orElseThrow().message());
  }

  @Test
  void doctorRequestReturnsStructuredInvalidReportForSemanticallyInvalidRequests()
      throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-invalid-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"OVERWRITE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, report.primaryProblem().orElseThrow().code());
    assertEquals("NEW", report.summary().orElseThrow().sourceType());
  }

  @Test
  void doctorRequestWritesProblemPointerToStderrWhenResponsePathCapturesInvalidReport()
      throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-invalid-response-");
    Path responsePath = Files.createTempFile("gridgrind-invalid-doctor-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--doctor-request",
                  "--execution-root",
                  workspace.toString(),
                  "--response",
                  responsePath.toString()
                },
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"OVERWRITE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    RequestDoctorReport report =
        GridGrindJson.readRequestDoctorReport(Files.readAllBytes(responsePath));

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(
        "GridGrind wrote the doctor report to "
            + responsePath.toAbsolutePath()
            + "; inspect that file for problems [INVALID_REQUEST: OVERWRITE persistence requires an"
            + " EXISTING source; a NEW workbook has no source file to overwrite]."
            + System.lineSeparator(),
        stderr.toString(StandardCharsets.UTF_8));
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, report.primaryProblem().orElseThrow().code());
  }

  @Test
  void doctorRequestPreflightsExistingWorkbookSourcesFromTheRequestFileDirectory()
      throws IOException {
    Path requestDirectory = Files.createTempDirectory("gridgrind-doctor-source-");
    Path requestPath = requestDirectory.resolve("doctor-existing.json");
    Files.writeString(
        requestPath,
        requestJson(
            "{ \"type\": \"EXISTING\", \"path\": \"missing-workbook.xlsx\" }",
            "{ \"type\": \"NONE\" }",
            """
            [
              {
                "stepId": "summary",
                "target": { "type": "WORKBOOK_CURRENT" },
                "query": { "type": "GET_WORKBOOK_SUMMARY" }
              }
            ]
            """));
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", requestPath.toString()},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    RequestDoctorReport report = doctorReport(stdout, stderr);

    assertEquals(1, exitCode);
    assertFalse(report.valid());
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND, report.primaryProblem().orElseThrow().code());
    assertEquals("OPEN_WORKBOOK", report.primaryProblem().orElseThrow().context().stage());
    assertEquals(
        java.util.Optional.of(requestDirectory.resolve("missing-workbook.xlsx").toString()),
        openWorkbookContext(report).sourceWorkbookPath());
  }

  @Test
  void doctorRequestReturnsCompactReadFailureWhenTheRequestFileCannotBeOpened() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    Path missingRequestPath =
        Path.of("tmp", "doctor-missing-" + UUID.randomUUID() + ".json").toAbsolutePath();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--request", missingRequestPath.toString()},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.code());
    assertEquals("doctor-request", failure.command());
    assertEquals(java.util.Optional.of("--request"), failure.argument());
    assertEquals("Request file not found: " + missingRequestPath, failure.message());
  }

  @Test
  void doctorRequestRequiresExecutionRootWhenRequestArrivesOnStdin() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request"},
                new ByteArrayInputStream(
                    requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.code());
    assertEquals("doctor-request", failure.command());
    assertEquals(java.util.Optional.of("--execution-root"), failure.argument());
    assertTrue(failure.message().contains("--execution-root"));
  }
}
