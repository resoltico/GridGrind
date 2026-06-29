package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import java.util.function.BiFunction;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;

/** Direct unit coverage for CLI doctor preflight normalization and binding selection. */
class CliDoctorRequestAnalyzerTest extends GridGrindCliTestSupport {
  @Test
  void diagnoseRejectsNonObjectRootsAndFallsBackToTemplateRequest() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                "[]".getBytes(StandardCharsets.UTF_8),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(1, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertEquals("NEW", sourceType(doctor.lastRequest()));
    assertEquals("NONE", persistenceType(doctor.lastRequest()));
    assertTrue(
        report.problems().stream()
            .map(GridGrindProblemDetail.Problem::message)
            .anyMatch("JSON request must be one object at the root"::equals));
  }

  @Test
  void diagnoseUsesRequestFileBindingsForCompleteAuthoredRequests() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    Path requestPath = Files.createTempFile("gridgrind-complete-request-", ".json");
    byte[] requestBytes =
        requestJson(
                "{ \"type\": \"EXISTING\", \"path\": \"input.xlsx\" }",
                "{ \"type\": \"SAVE_AS\", \"path\": \"output.xlsx\" }",
                "[]")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.of(requestPath),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertTrue(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(1, doctor.boundCalls());
    assertEquals(
        requestPath.getParent().toAbsolutePath().normalize(),
        doctor.lastInputs().workingDirectory());
    assertEquals(
        requestPath.getParent().toAbsolutePath().normalize().resolve(".gridgrind").resolve("tmp"),
        doctor.lastInputs().tempRoot());
    assertFalse(doctor.lastInputs().hasStandardInput());
    assertEquals("EXISTING", sourceType(doctor.lastRequest()));
    assertEquals("SAVE_AS", persistenceType(doctor.lastRequest()));
  }

  @Test
  void diagnoseUsesExplicitExecutionRootForStdinBackedRequests() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    Path workspace = Files.createTempDirectory("gridgrind-doctor-stdin-root-");
    byte[] requestBytes =
        requestJson(
                "{ \"type\": \"EXISTING\", \"path\": \"input.xlsx\" }",
                "{ \"type\": \"SAVE_AS\", \"path\": \"output.xlsx\" }",
                "[]")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.of(workspace),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertTrue(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(1, doctor.boundCalls());
    assertEquals(workspace.toAbsolutePath().normalize(), doctor.lastInputs().workingDirectory());
    assertEquals(
        workspace.toAbsolutePath().normalize().resolve(".gridgrind").resolve("tmp"),
        doctor.lastInputs().tempRoot());
    assertFalse(doctor.lastInputs().hasStandardInput());
  }

  @Test
  void diagnoseLeavesNewWorkbookSourcesUntouchedWhenExistingPathDefaultsDoNotApply()
      throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        requestJson(
                "{ \"type\": \"NEW\" }",
                "{ \"type\": \"SAVE_AS\", \"path\": \"output.xlsx\" }",
                "[]")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertTrue(report.valid());
    assertEquals(1, doctor.directCalls());
    assertEquals("NEW", sourceType(doctor.lastRequest()));
    assertEquals("", inputPath(doctor.lastRequest()));
  }

  @Test
  void diagnoseSynthesizesMissingDefaultsAndWorkbookPathsIntoOneBatch() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "EXISTING" },
          "persistence": { "type": "SAVE_AS" },
          "steps": []
        }
        """
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(1, doctor.directCalls());
    assertTrue(doctor.lastRequest().execution().isDefault());
    assertTrue(doctor.lastRequest().formulaEnvironment().isEmpty());
    assertEquals("__gridgrind_missing_source__.xlsx", inputPath(doctor.lastRequest()));
    assertEquals("__gridgrind_missing_output__.xlsx", outputPath(doctor.lastRequest()));
    List<String> problemMessages =
        report.problems().stream().map(GridGrindProblemDetail.Problem::message).toList();
    assertTrue(problemMessages.contains("Missing required field 'source.path'"));
    assertTrue(problemMessages.contains("Missing required field 'persistence.path'"));
    assertFalse(problemMessages.contains("Missing required field 'execution'"));
    assertFalse(problemMessages.contains("Missing required field 'formulaEnvironment'"));
  }

  @Test
  void diagnoseAcceptsMinimalRequestsAndAppliesTopLevelDefaults() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        minimalRequestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertTrue(report.valid());
    assertEquals(1, doctor.directCalls());
    assertTrue(doctor.lastRequest().execution().isDefault());
    assertTrue(doctor.lastRequest().formulaEnvironment().isEmpty());
  }

  @Test
  void diagnoseClassifiesMissingRootFieldsAsRequestShapeProblems() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
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
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(1, doctor.directCalls());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST_SHAPE, report.primaryProblem().orElseThrow().code());
    assertEquals(Optional.of("protocolVersion"), readRequestContext(report).jsonPath());
  }

  @Test
  void diagnoseRejectsExplicitNullTopLevelDefaultsBeforeDoctorExecution() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "EXISTING", "path": " " },
          "persistence": { "type": "SAVE_AS", "path": null },
          "execution": null,
          "formulaEnvironment": null,
          "steps": []
        }
        """
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertTrue(
        report.problems().stream()
            .map(GridGrindProblemDetail.Problem::message)
            .anyMatch(
                "Field 'execution' must be omitted when absent; explicit null is not accepted."
                    ::equals));
  }

  @Test
  void diagnoseTurnsTemplateCoveredExplicitNullsIntoFieldProblems() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        minimalRequestJson("{ \"type\": null }", "{ \"type\": \"NONE\" }", "[]")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(1, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertEquals("NEW", sourceType(doctor.lastRequest()));
    assertEquals(Optional.of("source.type"), readRequestContext(report).jsonPath());
    assertTrue(
        report.problems().stream()
            .map(GridGrindProblemDetail.Problem::message)
            .anyMatch(
                "Field 'source.type' must be omitted when absent; explicit null is not accepted."
                    ::equals));
  }

  @Test
  void diagnoseDoesNotTreatNumericWorkbookPathsAsMissing() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "EXISTING", "path": 123 },
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
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertFalse(
        report.problems().stream()
            .map(GridGrindProblemDetail.Problem::message)
            .anyMatch("Missing required field 'source.path'"::equals));
  }

  @Test
  void diagnoseDoesNotTreatBooleanSaveAsPathsAsMissing() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "NEW" },
          "persistence": { "type": "SAVE_AS", "path": true },
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
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertFalse(
        report.problems().stream()
            .map(GridGrindProblemDetail.Problem::message)
            .anyMatch("Missing required field 'persistence.path'"::equals));
  }

  @Test
  void diagnoseTreatsObjectWorkbookPathsAsMissingConditionalPaths() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "EXISTING", "path": {} },
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
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(1, doctor.directCalls());
    assertEquals("__gridgrind_missing_source__.xlsx", inputPath(doctor.lastRequest()));
    assertTrue(
        report.problems().stream()
            .map(GridGrindProblemDetail.Problem::message)
            .anyMatch("Missing required field 'source.path'"::equals));
  }

  @Test
  void diagnoseDeduplicatesDoctorProblemsAgainstPreflightProblems() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor(
            (request, inputs) ->
                RequestDoctorReport.invalid(
                    summaryFor(request), List.of(), missingFieldProblem("source.path")));
    byte[] requestBytes =
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "EXISTING" },
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
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(
        1,
        report.problems().stream()
            .filter(problem -> "Missing required field 'source.path'".equals(problem.message()))
            .count());
  }

  @Test
  void diagnoseRejectsScalarWorkbookSourceShapesBeforeDoctorExecution() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        requestJson("\"EXISTING\"", "{ \"type\": \"NONE\" }", "[]")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST_SHAPE, report.primaryProblem().orElseThrow().code());
  }

  @Test
  void diagnoseRejectsObjectTypedScalarFieldsAndMissingWorkbookDiscriminators() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        requestJson("{ }", "{ \"type\": \"NONE\" }", "[]")
            .replace("\"protocolVersion\": \"V1\"", "\"protocolVersion\": {}")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST_SHAPE, report.primaryProblem().orElseThrow().code());
  }

  @Test
  void diagnoseCarriesExistingRequestShapeForStandardInputBindingConflicts() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        requestJson(
                "{ \"type\": \"EXISTING\", \"path\": \"input.xlsx\" }",
                "{ \"type\": \"OVERWRITE\" }",
                """
                [
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
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals("EXISTING", report.summary().orElseThrow().sourceType());
    assertEquals("OVERWRITE", report.summary().orElseThrow().persistenceType());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, report.primaryProblem().orElseThrow().code());
  }

  @Test
  void diagnoseCarriesSaveAsRequestShapeForStandardInputBindingConflicts() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        requestJson(
                "{ \"type\": \"EXISTING\", \"path\": \"input.xlsx\" }",
                "{ \"type\": \"SAVE_AS\", \"path\": \"output.xlsx\" }",
                """
                [
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
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                requestBytes,
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals("EXISTING", report.summary().orElseThrow().sourceType());
    assertEquals("SAVE_AS", report.summary().orElseThrow().persistenceType());
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, report.primaryProblem().orElseThrow().code());
  }

  private static RequestDoctorReport.Summary summaryFor(WorkbookPlan request) {
    int stepCount = request.steps().size();
    return new RequestDoctorReport.Summary(
        sourceType(request),
        persistenceType(request),
        request.effectiveExecutionMode().modeType(),
        request.calculationPolicy().effectiveStrategy().strategyType(),
        request.calculationPolicy().markRecalculateOnOpen(),
        false,
        stepCount,
        stepCount,
        0,
        0);
  }

  private static GridGrindProblemDetail.Problem missingFieldProblem(String jsonPath) {
    return GridGrindProblems.fromException(
        new InvalidRequestShapeException(
            GridGrindRequestProblemSupport.missingRequiredFieldMessage(jsonPath),
            Optional.of(jsonPath),
            Optional.empty(),
            Optional.empty(),
            null),
        new ProblemContext.ReadRequest(
            ProblemContextRequestSurfaces.RequestInput.standardInput(),
            ProblemContextRequestSurfaces.JsonLocation.pathOnly(jsonPath)));
  }

  private static String sourceType(WorkbookPlan request) {
    return requiredString(GridGrindJson.requestTree(request).path("source").path("type"));
  }

  private static String persistenceType(WorkbookPlan request) {
    return requiredString(GridGrindJson.requestTree(request).path("persistence").path("type"));
  }

  private static String inputPath(WorkbookPlan request) {
    return optionalString(GridGrindJson.requestTree(request).path("source").path("path"));
  }

  private static String outputPath(WorkbookPlan request) {
    return optionalString(GridGrindJson.requestTree(request).path("persistence").path("path"));
  }

  private static String requiredString(JsonNode node) {
    return node.stringValue();
  }

  private static String optionalString(JsonNode node) {
    return node.isString() ? node.stringValue() : "";
  }

  /** Test double that records which doctor entry point the analyzer selected. */
  private static final class RecordingDoctor implements GridGrindRequestDoctor {
    private final BiFunction<WorkbookPlan, Optional<GridGrindRequestInputs>, RequestDoctorReport>
        responder;
    private int directCalls;
    private int boundCalls;
    private WorkbookPlan lastRequest;
    private GridGrindRequestInputs lastInputs;

    private RecordingDoctor(
        BiFunction<WorkbookPlan, Optional<GridGrindRequestInputs>, RequestDoctorReport> responder) {
      this.responder = responder;
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request) {
      directCalls++;
      lastRequest = request;
      return responder.apply(request, Optional.empty());
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request, GridGrindRequestInputs inputs) {
      boundCalls++;
      lastRequest = request;
      lastInputs = inputs;
      return responder.apply(request, Optional.of(inputs));
    }

    private int directCalls() {
      return directCalls;
    }

    private int boundCalls() {
      return boundCalls;
    }

    private WorkbookPlan lastRequest() {
      return lastRequest;
    }

    private GridGrindRequestInputs lastInputs() {
      return lastInputs;
    }
  }
}
