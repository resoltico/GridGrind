package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
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
                "shape", Optional.of("steps[0]"), Optional.of(3), Optional.of(7), null));

    assertEquals(
        List.of(
            "gridgrind --print-protocol-catalog --search <term>",
            "gridgrind --print-request-template --response request.json"),
        failure.suggestions());
    assertEquals(Optional.of("steps[0]"), failure.location().jsonPath());
    assertEquals(Optional.of(3), failure.location().jsonLine());
    assertEquals(Optional.of(7), failure.location().jsonColumn());
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
    return GridGrindProblemDetail.Problem.of(
        code,
        message,
        new ProblemContext.ReadRequest(RequestInput.standardInput(), JsonLocation.unavailable()));
  }
}
