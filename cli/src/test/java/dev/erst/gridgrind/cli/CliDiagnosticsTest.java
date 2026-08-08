package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.MessageShape;
import dev.erst.gridgrind.contract.json.MissingRequiredField;
import dev.erst.gridgrind.contract.json.MissingTypeDiscriminator;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for canonical CLI diagnostic branching. */
class CliDiagnosticsTest extends GridGrindCliTestSupport {
  @Test
  void readRequestFailureUsesDoctorSuggestionsForDoctorRequestShapeErrors() {
    InvalidRequestShapeException exception =
        new InvalidRequestShapeException(
            new MessageShape("shape", Optional.of("steps[0]")),
            Optional.of("steps[0]"),
            Optional.of(3),
            Optional.of(7),
            null);
    CliDiagnostic failure =
        CliDiagnostics.readRequestFailure(
            1,
            "doctor-request",
            GridGrindProblems.fromException(
                exception,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.unavailable())));

    assertEquals(
        List.of(
            "gridgrind --print-protocol-catalog --search \"steps\"",
            "gridgrind --print-request-template --response request.json"),
        failure.suggestions());
    assertEquals(Optional.of("steps[0]"), readRequestContext(failure).jsonPath());
    assertEquals(Optional.of(3), readRequestContext(failure).jsonLine());
    assertEquals(Optional.of(7), readRequestContext(failure).jsonColumn());
  }

  @Test
  void readRequestFailureFallsBackToGenericGuidanceForUnhandledCodes() {
    CliDiagnostic failure =
        CliDiagnostics.readRequestFailure(
            1, "execute", problem(GridGrindProblemCode.WORKBOOK_NOT_FOUND, "missing workbook"));

    assertEquals(List.of("gridgrind --help", "gridgrind --help-protocol"), failure.suggestions());
    assertEquals(
        "Create the workbook first or provide an existing workbook path.",
        failure.problem().resolution());
  }

  @Test
  void readRequestFailureExplainsDoctorInvariantRepairs() {
    CliDiagnostic failure =
        CliDiagnostics.readRequestFailure(
            1, "doctor-request", problem(GridGrindProblemCode.INVALID_REQUEST, "invalid request"));

    assertEquals("Fix the request data and retry the workflow.", failure.problem().resolution());
  }

  @Test
  void readRequestFailureUsesSpecificResolutionForMissingRequiredFields() {
    InvalidRequestShapeException failureCause =
        new InvalidRequestShapeException(
            new MissingRequiredField("protocolVersion"),
            Optional.of("protocolVersion"),
            Optional.empty(),
            Optional.empty(),
            null);
    CliDiagnostic failure =
        CliDiagnostics.readRequestFailure(
            1,
            "execute",
            GridGrindProblems.fromException(
                failureCause,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.pathOnly("protocolVersion"))));

    assertEquals(Optional.of("protocolVersion"), readRequestContext(failure).jsonPath());
    assertEquals(
        "Add protocolVersion: \"V2\" at the request root.", failure.problem().resolution());
  }

  @Test
  void readRequestFailureQualifiesLocalDiscriminatorPathsFromTheRequestContext() {
    InvalidRequestShapeException failureCause =
        new InvalidRequestShapeException(
            new MissingTypeDiscriminator("type"),
            Optional.of("steps[0].target.type"),
            Optional.empty(),
            Optional.empty(),
            null);
    CliDiagnostic failure =
        CliDiagnostics.readRequestFailure(
            1,
            "execute",
            GridGrindProblems.fromException(
                failureCause,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.unavailable())));

    assertEquals(Optional.of("steps[0].target.type"), readRequestContext(failure).jsonPath());
    assertEquals(
        "Add the required type discriminator at 'steps[0].target.type'.",
        failure.problem().resolution());
  }

  @Test
  void readRequestFailureFallsBackToContextLocationWhenExceptionCarriesNone() {
    CliDiagnostic failure =
        CliDiagnostics.readRequestFailure(
            1,
            "execute",
            problem(
                GridGrindProblemCode.INVALID_REQUEST,
                "invalid request",
                JsonLocation.pathOnly("steps[1].stepId")));

    assertEquals(Optional.of("steps[1].stepId"), readRequestContext(failure).jsonPath());
    assertEquals("Fix the request data and retry the workflow.", failure.problem().resolution());
  }

  @Test
  void responseWriteFailureOmitsBlankStdoutSuggestionHints() {
    CliDiagnostic failure =
        CliDiagnostics.responseWriteFailure(
            "print-request-template",
            "request template",
            Path.of("/tmp/response.json"),
            new java.io.IOException("denied"),
            Optional.of("  "));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals(List.of(), failure.suggestions());
    assertEquals(Optional.of("/tmp/response.json"), writeResponseContext(failure).responsePath());
  }

  @Test
  void readRequestFailuresRejectsAnEmptyProblemCollection() {
    assertEquals(
        "problems must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> CliDiagnostics.readRequestFailures(1, "execute", List.of()))
            .getMessage());
  }

  private static GridGrindProblemDetail.Problem problem(GridGrindProblemCode code, String message) {
    return problem(code, message, JsonLocation.unavailable());
  }

  private static GridGrindProblemDetail.Problem problem(
      GridGrindProblemCode code, String message, JsonLocation jsonLocation) {
    return GridGrindProblemDetail.Problem.of(
        code, message, new ProblemContext.ReadRequest(RequestInput.standardInput(), jsonLocation));
  }
}
