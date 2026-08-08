package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.GridGrindRecipeCatalog;
import dev.erst.gridgrind.cli.discovery.RecipeCatalog;
import dev.erst.gridgrind.cli.discovery.RecipeKeywordMatchReport;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipe;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
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
import org.junit.jupiter.api.Test;

/** Catalog and discovery command integration tests for GridGrindCli. */
class GridGrindCliCatalogCommandTest extends GridGrindCliTestSupport {
  @Test
  void printRecipeFlagPrintsKnownGeneratedExampleAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe", "--lookup", "WORKBOOK_HEALTH"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());

    assertEquals(0, exitCode);
    assertEquals(GridGrindShippedExamples.find("WORKBOOK_HEALTH").orElseThrow().plan(), request);
    assertFalse(stdout.toString(StandardCharsets.UTF_8).contains(": null"));
  }

  @Test
  void printRecipeCatalogFlagPrintsUnifiedRecipeCatalogAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-catalog"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    RecipeCatalog catalog = GridGrindCliJson.readBytes(stdout.toByteArray(), RecipeCatalog.class);

    assertEquals(0, exitCode);
    assertEquals(GridGrindRecipeCatalog.catalog(), catalog);
    assertEquals(
        dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode.SELF_CONTAINED,
        catalog.recipes().stream()
            .filter(recipe -> "BUDGET".equals(recipe.id()))
            .findFirst()
            .orElseThrow()
            .workspaceMode());
    assertEquals(
        dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
        catalog.recipes().stream()
            .filter(recipe -> "PACKAGE_SECURITY_INSPECTION".equals(recipe.id()))
            .findFirst()
            .orElseThrow()
            .workspaceMode());
    assertEquals(
        java.util.List.of("package-security-assets/gridgrind-package-security.xlsx"),
        catalog.recipes().stream()
            .filter(recipe -> "PACKAGE_SECURITY_INSPECTION".equals(recipe.id()))
            .findFirst()
            .orElseThrow()
            .requiredWorkspacePaths());
  }

  @Test
  void printAssetBackedExampleWarnsOnStderrBeforeExecution() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe", "--lookup", "PACKAGE_SECURITY_INSPECTION"},
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
  void printRecipeFlagRejectsUnknownExampleId() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe", "--lookup", "BOGUS_EXAMPLE"},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stderr);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--lookup"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("BOGUS_EXAMPLE"));
    assertTrue(failure.problem().message().contains("--print-recipe-catalog"));
  }

  @Test
  void printRecipeFlagWritesUnknownExampleFailureToResponsePathAndPointsStderrAtIt()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-unknown-example-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-recipe",
                  "--lookup",
                  "BOGUS_EXAMPLE",
                  "--response",
                  responsePath.toString()
                },
                InputStream.nullInputStream(),
                stdout,
                stderr);

    CliDiagnostic failure = cliDiagnostic(Files.readAllBytes(responsePath));
    CliDiagnostic stderrDiagnostic = cliDiagnosticOnStderr(stderr);

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(failure, stderrDiagnostic);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertTrue(failure.problem().message().contains("BOGUS_EXAMPLE"));
  }

  @Test
  void printRecipeFlagSuggestsStableIdForCommonNonCanonicalTokens() throws IOException {
    assertSuggestedExampleId("chart");
    assertSuggestedExampleId("chart-request.json");
    assertSuggestedExampleId("chart-request");
    assertSuggestedExampleId("chart request");
  }

  @Test
  void printRecipeCatalogContainsPublishedTaskStarterMetadata() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-catalog"},
                new ByteArrayInputStream("ignored".getBytes(StandardCharsets.UTF_8)),
                stdout);

    RecipeCatalog catalog = GridGrindCliJson.readBytes(stdout.toByteArray(), RecipeCatalog.class);

    assertEquals(0, exitCode);
    assertTrue(catalog.recipes().stream().anyMatch(recipe -> "DASHBOARD".equals(recipe.id())));
    assertTrue(
        catalog.recipes().stream()
            .filter(recipe -> "DASHBOARD".equals(recipe.id()))
            .findFirst()
            .orElseThrow()
            .summary()
            .contains("dashboard"));
  }

  private static void assertSuggestedExampleId(String authoredValue) throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe", "--lookup", authoredValue},
                new ByteArrayInputStream(new byte[0]),
                stdout,
                stdout); // stdout passed for both streams; use the bytes overload directly
    CliDiagnostic failure = cliDiagnostic(stdout.toByteArray());
    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--lookup"), parseArgumentsContext(failure).argumentName());
    assertTrue(
        failure.problem().message().contains("did you mean CHART?"),
        () -> "expected CHART suggestion for authored value " + authoredValue);
    assertTrue(
        failure.problem().message().contains("--print-recipe-catalog"),
        () -> "expected recovery guidance for authored value " + authoredValue);
  }

  @Test
  void printRecipeCatalogWithRecipeFilterReturnsMatchingEntry() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-catalog", "--lookup", "DASHBOARD"},
                InputStream.nullInputStream(),
                stdout);

    assertEquals(0, exitCode);
    assertTrue(
        stdout.toString(StandardCharsets.UTF_8).contains("\"DASHBOARD\""),
        "output must contain the task id");
  }

  @Test
  void printRecipeCatalogRejectsBlankRecipeFilterWithStructuredFailure() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-catalog", "--lookup", ""},
                InputStream.nullInputStream(),
                stdout,
                stderr);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("print-recipe-catalog", failure.command());
    assertEquals(java.util.Optional.of("--lookup"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("recipe lookup id must not be blank"));
  }

  @Test
  void printRecipeCatalogWithUnknownRecipeReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-catalog", "--lookup", "BOGUS_TASK"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--lookup"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("BOGUS_TASK"));
    assertTrue(failure.problem().message().contains("--print-recipe-catalog"));
    assertTrue(
        failure
            .problem()
            .message()
            .contains("--print-recipe-keyword-match --query \"monthly sales dashboard\""));
  }

  @Test
  void printRecipeCatalogWithNonCanonicalRecipeIdSuggestsStableToken() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-catalog", "--lookup", "dashboard"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--lookup"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("did you mean DASHBOARD?"));
    assertTrue(failure.problem().message().contains("--print-recipe-catalog"));
    assertTrue(
        failure
            .problem()
            .message()
            .contains("--print-recipe-keyword-match --query \"monthly sales dashboard\""));
  }

  @Test
  void printRecipeCatalogWithNormalizedRecipeIdSuggestsStableToken() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-catalog", "--lookup", "audit existing workbook"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--lookup"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("did you mean AUDIT_EXISTING_WORKBOOK?"));
    assertTrue(failure.problem().message().contains("--print-recipe-catalog"));
    assertTrue(
        failure
            .problem()
            .message()
            .contains("--print-recipe-keyword-match --query \"monthly sales dashboard\""));
  }

  @Test
  void printRecipeCatalogRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-catalog", "--version"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--version"), parseArgumentsContext(failure).argumentName());
    assertTrue(
        failure
            .problem()
            .message()
            .contains(
                "Only one primary command may be used per invocation; --print-recipe-catalog"
                    + " cannot be combined with --version"));
  }

  @Test
  void printRecipeFlagPrintsRunnableStarterRequestAndReturnsExitCodeZero() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe", "--lookup", "DASHBOARD"},
                InputStream.nullInputStream(),
                stdout);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());
    String rendered = stdout.toString(StandardCharsets.UTF_8);
    GridGrindCliRecipe recipe =
        GridGrindCliRecipeRegistry.recipeFor("DASHBOARD")
            .filter(
                candidate ->
                    candidate.view() == dev.erst.gridgrind.cli.discovery.RecipeView.TASK_STARTER)
            .orElseThrow();

    assertEquals(0, exitCode);
    assertEquals(recipe.plan(), request);
    assertTrue(request.execution().isDefault());
    assertTrue(request.formulaEnvironment().isEmpty());
    assertFalse(rendered.contains("\"execution\""));
    assertFalse(rendered.contains("\"formulaEnvironment\""));
  }

  @Test
  void printAssetBackedTaskStarterWarnsOnStderrBeforeExecution() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe", "--lookup", "AUDIT_EXISTING_WORKBOOK"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    WorkbookPlan request = GridGrindJson.readRequest(stdout.toByteArray());
    GridGrindCliRecipe recipe =
        GridGrindCliRecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK")
            .filter(
                candidate ->
                    candidate.view() == dev.erst.gridgrind.cli.discovery.RecipeView.TASK_STARTER)
            .orElseThrow();

    assertEquals(0, exitCode);
    assertEquals(recipe.plan(), request);
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
  void printRecipeWithUnknownTaskReturnsError() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe", "--lookup", "BOGUS_TASK"},
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--lookup"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("BOGUS_TASK"));
    assertTrue(failure.problem().message().contains("--print-recipe-catalog"));
    assertTrue(
        failure
            .problem()
            .message()
            .contains("--print-recipe-keyword-match --query \"monthly sales dashboard\""));
  }

  @Test
  void printRecipeRejectsTrailingExecutionFlags() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-recipe", "--lookup", "DASHBOARD", "--request", "ignored.json"
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
  void printRecipeKeywordMatchFlagPrintsRankedRecipeMatchesAndReturnsExitCodeZero()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-recipe-keyword-match", "--query", "monthly sales dashboard with charts"
                },
                InputStream.nullInputStream(),
                stdout);

    RecipeKeywordMatchReport report =
        GridGrindCliJson.readBytes(stdout.toByteArray(), RecipeKeywordMatchReport.class);

    assertEquals(0, exitCode);
    assertEquals("monthly sales dashboard with charts", report.query());
    assertEquals("DASHBOARD", report.candidates().getFirst().recipeId());
    assertTrue(report.candidates().getFirst().matchedTerms().contains("dashboard"));
    assertTrue(report.candidates().getFirst().matchedTerms().contains("chart"));
  }

  @Test
  void printTaskKeywordMatchRejectsBlankQueries() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-keyword-match", "--query", ""},
                InputStream.nullInputStream(),
                stdout,
                stderr);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--query"), parseArgumentsContext(failure).argumentName());
  }

  @Test
  void printTaskKeywordMatchRejectsQueriesThatNormalizeToNoSearchableTerms() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--print-recipe-keyword-match", "--query", "a"},
                InputStream.nullInputStream(),
                stdout,
                stderr);
    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals(java.util.Optional.of("--query"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("searchable term after normalization"));
  }

  @Test
  void printRequestTemplatePrintsTheCurrentMachineReadableTemplate() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--print-request-template"}, InputStream.nullInputStream(), stdout);

    dev.erst.gridgrind.contract.dto.WorkbookPlan template =
        GridGrindJson.readRequest(stdout.toByteArray());
    String rendered = stdout.toString(StandardCharsets.UTF_8);

    assertEquals(0, exitCode);
    assertEquals("FULL_XSSF", template.execution().mode().modeType());
    assertTrue(template.formulaEnvironment().isEmpty());
    assertFalse(rendered.contains("\"execution\""));
    assertFalse(rendered.contains("\"formulaEnvironment\""));
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
    assertEquals(java.util.Optional.empty(), readRequestContext(report).jsonLine());
    assertEquals(java.util.Optional.empty(), readRequestContext(report).jsonColumn());
  }

  @Test
  void doctorRequestRejectsMissingRequestInputWithCompactCliDiagnostic() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(new String[] {"--doctor-request"}, InputStream.nullInputStream(), stdout, stderr);

    CliDiagnostic failure = cliDiagnosticOnStderr(stdout, stderr);

    assertEquals(2, exitCode);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, failure.problem().code());
    assertEquals("doctor-request", failure.command());
    assertEquals(java.util.Optional.of("--request"), parseArgumentsContext(failure).argumentName());
    assertTrue(failure.problem().message().contains("No request JSON was provided."));
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
                      "protocolVersion": "V2",
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
    assertTrue(problemMessages.contains("Missing required field 'persistence.ifExists'"));
  }

  @Test
  void doctorRequestBatchesIndependentSemanticValidationProblems() throws IOException {
    Path workspace = Files.createTempDirectory("gridgrind-doctor-semantic-batch-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--doctor-request", "--execution-root", workspace.toString()},
                new ByteArrayInputStream(
                    """
                    {
                      "protocolVersion": "V2",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "OVERWRITE" },
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
                      "steps": []
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
    assertEquals(2, report.problems().size());
    assertTrue(
        problemMessages.contains(
            "execution.mode.type=EVENT_READ requires"
                + " execution.calculation.strategy=DO_NOT_CALCULATE and"
                + " markRecalculateOnOpen=false"));
    assertTrue(
        problemMessages.contains(
            "OVERWRITE persistence requires an EXISTING source; a NEW workbook has no source"
                + " file to overwrite"));
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
}
