package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
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

/** Direct coverage for tolerant doctor intake and complete-plan binding selection. */
class CliDoctorRequestAnalyzerTest extends GridGrindCliTestSupport {
  @Test
  void diagnoseRunsTheDoctorForACompleteMinimalV2Request() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis(
                    minimalRequestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                InputStream.nullInputStream());

    assertTrue(report.valid());
    assertEquals(1, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertEquals("NEW", report.summary().orElseThrow().sourceType());
  }

  @Test
  void diagnoseUsesRequestFileBindingsForACompleteRequest() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    Path requestPath = Files.createTempFile("gridgrind-complete-request-", ".json");
    byte[] requestBytes =
        requestJson(
                "{ \"type\": \"EXISTING\", \"path\": \"input.xlsx\" }",
                "{ \"type\": \"SAVE_AS\", \"path\": \"output.xlsx\", \"ifExists\": \"REJECT\" }",
                "[]")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.of(requestPath),
                Optional.empty(),
                Optional.empty(),
                analysis(requestBytes),
                InputStream.nullInputStream());

    assertTrue(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(1, doctor.boundCalls());
    assertEquals(
        requestPath.getParent().toAbsolutePath().normalize(),
        doctor.lastInputs().orElseThrow().workingDirectory());
  }

  @Test
  void diagnoseReportsIndependentStructuralFaultsWithoutSyntheticBinding() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW", "unexpected": true },
          "persistence": null,
          "steps": [
            {
              "stepId": 7,
              "target": { "type": "WORKBOOK_CURRENT" },
              "action": { "type": "CLEAR_WORKBOOK_PROTECTION" }
            }
          ]
        }
        """
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis(requestBytes),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    List<String> messages =
        report.problems().stream().map(GridGrindProblemDetail.Problem::message).toList();
    assertTrue(messages.contains("Unknown field 'source.unexpected'"));
    assertTrue(
        messages.contains(
            "Field 'persistence' must be omitted when absent; explicit null is not accepted."));
    assertTrue(messages.contains("Field 'steps[0].stepId' must be a JSON string"));
  }

  @Test
  void diagnoseRejectsV1WithoutTranslatingIt() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));
    byte[] requestBytes =
        minimalRequestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
            .replace("\"protocolVersion\": \"V2\"", "\"protocolVersion\": \"V1\"")
            .getBytes(StandardCharsets.UTF_8);

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis(requestBytes),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertTrue(
        report.problems().stream()
            .map(GridGrindProblemDetail.Problem::message)
            .anyMatch(message -> message.contains("protocolVersion") && message.contains("V2")));
  }

  @Test
  void diagnoseClassifiesMalformedUtf8AsInvalidEncoding() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis(new byte[] {'{', '"', 'x', '"', ':', (byte) 0xC3, (byte) 0x28}),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(
        GridGrindProblemCode.INVALID_ENCODING, report.primaryProblem().orElseThrow().code());
  }

  @Test
  void diagnoseDoesNotFabricateAPlanForANonObjectRoot() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis("[]".getBytes(StandardCharsets.UTF_8)),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertEquals(
        "Field 'request' must be a JSON object at the root",
        report.primaryProblem().orElseThrow().message());
  }

  @Test
  void diagnoseReportsConstructorValidationFailuresWithoutReparsingTheRequest() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis(
                    minimalRequestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            """
                        [
                          {
                            "stepId": "duplicate",
                            "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                            "action": { "type": "ENSURE_SHEET" }
                          },
                          {
                            "stepId": "duplicate",
                            "target": { "type": "SHEET_BY_NAME", "name": "Forecast" },
                            "action": { "type": "ENSURE_SHEET" }
                          }
                        ]
                            """)
                        .getBytes(StandardCharsets.UTF_8)),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertEquals(
        "steps must not contain duplicate stepId values: duplicate",
        report.primaryProblem().orElseThrow().message());
  }

  @Test
  void diagnoseBatchesEveryIndependentConstructorFailureFromOneAnalysis() throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis(
                    minimalRequestJson(
                            "{ \"type\": \"EXISTING\", \"path\": \"input.xls\" }",
                            "{ \"type\": \"SAVE_AS\", \"path\": \"output.txt\", \"ifExists\": \"REJECT\" }",
                            "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(0, doctor.directCalls());
    assertEquals(0, doctor.boundCalls());
    assertEquals(
        List.of("source.path", "persistence.path"),
        report.problems().stream()
            .map(problem -> assertInstanceOf(ProblemContext.ReadRequest.class, problem.context()))
            .map(context -> context.jsonPath().orElseThrow())
            .toList());
    assertEquals(
        List.of("path must end in .xlsx (got: '.xls')", "path must end in .xlsx (got: '.txt')"),
        report.problems().stream().map(GridGrindProblemDetail.Problem::message).toList());
  }

  @Test
  void diagnoseRetainsDoctorProblemsWithTheStandardInputBindingProblem() throws IOException {
    RecordingDoctor doctor = new RecordingDoctor((request, inputs) -> invalidReportFor(request));

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis(
                    minimalRequestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"SAVE_AS\", \"path\": \"output.xlsx\", \"ifExists\": \"REPLACE\" }",
                            standardInputSteps())
                        .getBytes(StandardCharsets.UTF_8)),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(1, doctor.directCalls());
    assertEquals(2, report.problems().size());
    assertTrue(report.problems().getFirst().message().contains("STANDARD_INPUT"));
    assertEquals("Doctor preflight failed", report.problems().get(1).message());
  }

  @Test
  void diagnoseDescribesExistingOverwriteRequestsWhenStandardInputCannotBeBound()
      throws IOException {
    RecordingDoctor doctor =
        new RecordingDoctor((request, inputs) -> RequestDoctorReport.clean(summaryFor(request)));

    RequestDoctorReport report =
        new CliDoctorRequestAnalyzer(doctor)
            .diagnose(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                analysis(
                    minimalRequestJson(
                            "{ \"type\": \"EXISTING\", \"path\": \"input.xlsx\" }",
                            "{ \"type\": \"OVERWRITE\" }",
                            standardInputSteps())
                        .getBytes(StandardCharsets.UTF_8)),
                InputStream.nullInputStream());

    assertFalse(report.valid());
    assertEquals(1, doctor.directCalls());
    assertEquals("EXISTING", report.summary().orElseThrow().sourceType());
    assertEquals("OVERWRITE", report.summary().orElseThrow().persistenceType());
  }

  private static RequestDoctorReport invalidReportFor(WorkbookPlan request) {
    return RequestDoctorReport.invalid(
        summaryFor(request),
        List.of(),
        GridGrindProblems.problem(
            GridGrindProblemCode.INVALID_REQUEST,
            "Doctor preflight failed",
            new ProblemContext.ValidateRequest(
                ProblemContextRequestSurfaces.RequestShape.unknown()),
            List.of()));
  }

  private static RequestAnalysis analysis(byte[] requestBytes) {
    return GridGrindJson.analyzeRequest(requestBytes);
  }

  private static String standardInputSteps() {
    return """
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
        """;
  }

  private static RequestDoctorReport.Summary summaryFor(WorkbookPlan request) {
    int stepCount = request.steps().size();
    return new RequestDoctorReport.Summary(
        switch (request.source()) {
          case WorkbookPlan.WorkbookSource.New _ -> "NEW";
          case WorkbookPlan.WorkbookSource.ExistingFile _ -> "EXISTING";
        },
        switch (request.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
          case WorkbookPlan.WorkbookPersistence.Overwrite _ -> "OVERWRITE";
          case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
        },
        request.effectiveExecutionMode().modeType(),
        request.calculationPolicy().effectiveStrategy().strategyType(),
        request.calculationPolicy().markRecalculateOnOpen(),
        false,
        stepCount,
        stepCount,
        0,
        0);
  }

  /** Test double that records which doctor entry point the analyzer selected. */
  private static final class RecordingDoctor implements GridGrindRequestDoctor {
    private final BiFunction<WorkbookPlan, Optional<GridGrindRequestInputs>, RequestDoctorReport>
        responder;
    private int directCalls;
    private int boundCalls;
    private Optional<GridGrindRequestInputs> lastInputs = Optional.empty();

    private RecordingDoctor(
        BiFunction<WorkbookPlan, Optional<GridGrindRequestInputs>, RequestDoctorReport> responder) {
      this.responder = responder;
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request) {
      directCalls++;
      return responder.apply(request, Optional.empty());
    }

    @Override
    public RequestDoctorReport diagnose(WorkbookPlan request, GridGrindRequestInputs inputs) {
      boundCalls++;
      lastInputs = Optional.of(inputs);
      return responder.apply(request, lastInputs);
    }

    private int directCalls() {
      return directCalls;
    }

    private int boundCalls() {
      return boundCalls;
    }

    private Optional<GridGrindRequestInputs> lastInputs() {
      return lastInputs;
    }
  }
}
