package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
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

/** Focused coverage for compact CLI failure-report branching. */
class CliFailureReportsTest {
  @Test
  void readRequestFailureUsesDoctorSuggestionsForDoctorRequestShapeErrors() {
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "doctor-request",
            Optional.of("--request"),
            problem(GridGrindProblemCode.INVALID_REQUEST_SHAPE, "shape"),
            new InvalidRequestShapeException(
                new MessageShape("shape", Optional.of("steps[0]")),
                Optional.of("steps[0]"),
                Optional.of(3),
                Optional.of(7),
                null));

    assertEquals(
        List.of(
            "gridgrind --print-protocol-catalog --search \"sheet layout\"",
            "gridgrind --print-request-template --response request.json"),
        failure.suggestions());
    assertEquals(Optional.of("steps[0]"), failure.location().orElseThrow().jsonPath());
    assertEquals(Optional.of(3), failure.location().orElseThrow().jsonLine());
    assertEquals(Optional.of(7), failure.location().orElseThrow().jsonColumn());
  }

  @Test
  void readRequestFailureFallsBackToGenericGuidanceForUnhandledCodes() {
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "execute",
            Optional.of("--request"),
            problem(GridGrindProblemCode.WORKBOOK_NOT_FOUND, "missing workbook"),
            new RuntimeException("missing workbook"));

    assertEquals(List.of("gridgrind --help", "gridgrind --help-protocol"), failure.suggestions());
    assertEquals(
        Optional.of("Correct the request input, then rerun the command."), failure.resolution());
  }

  @Test
  void readRequestFailureExplainsDoctorInvariantRepairs() {
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "doctor-request",
            Optional.of("--request"),
            problem(GridGrindProblemCode.INVALID_REQUEST, "invalid request"),
            new RuntimeException("invalid request"));

    assertEquals(
        Optional.of(
            "The request document decoded but violates request invariants. Correct the authored"
                + " request and rerun --doctor-request."),
        failure.resolution());
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
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "execute",
            Optional.of("--request"),
            GridGrindProblems.fromException(
                failureCause,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.pathOnly("protocolVersion"))),
            failureCause);

    assertEquals(Optional.of("protocolVersion"), failure.location().orElseThrow().jsonPath());
    assertEquals(
        Optional.of(
            "Add protocolVersion: \"V1\" at the request root. Use"
                + " --print-protocol-catalog --search \"sheet layout\" or --help-protocol when"
                + " you need the authoritative field and discriminator contract."),
        failure.resolution());
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
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "execute",
            Optional.of("--request"),
            GridGrindProblems.fromException(
                failureCause,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.unavailable())),
            failureCause);

    assertEquals(Optional.of("steps[0].target.type"), failure.location().orElseThrow().jsonPath());
    assertEquals(
        Optional.of(
            "Add the required type discriminator at 'steps[0].target.type'. Use"
                + " --print-protocol-catalog --search \"sheet layout\" or --help-protocol when"
                + " you need the authoritative field and discriminator contract."),
        failure.resolution());
  }

  @Test
  void readRequestFailureFallsBackToContextLocationWhenExceptionCarriesNone() {
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "execute",
            Optional.of("--request"),
            problem(
                GridGrindProblemCode.INVALID_REQUEST,
                "invalid request",
                JsonLocation.pathOnly("steps[1].stepId")),
            new RuntimeException("invalid request"));

    assertEquals(Optional.of("steps[1].stepId"), failure.location().orElseThrow().jsonPath());
    assertEquals(
        Optional.of(
            "The request document decoded but violates request invariants. Run --doctor-request"
                + " first, then execute the corrected request."),
        failure.resolution());
  }

  @Test
  void responseWriteFailureOmitsBlankStdoutSuggestionHints() {
    CliFailureReport failure =
        CliFailureReports.responseWriteFailure(
            "print-request-template",
            "request template",
            Path.of("/tmp/response.json"),
            new java.io.IOException("denied"),
            Optional.of("  "));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.code());
    assertEquals(List.of(), failure.suggestions());
    assertEquals(Optional.of("--response"), failure.argument());
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
