package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Residual coverage for compact CLI failure-report branches. */
class CliFailureReportsResidualTest {
  @Test
  void readRequestFailureUsesExecutionRepairTextForGenericInvalidRequests() {
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "execute",
            Optional.of("--request"),
            problem(GridGrindProblemCode.INVALID_REQUEST, "invalid request"),
            new RuntimeException("invalid request"));

    assertEquals(
        Optional.of(
            "The request document decoded but violates request invariants. Run --doctor-request"
                + " first, then execute the corrected request."),
        failure.resolution());
  }

  @Test
  void readRequestFailureUsesDoctorRepairTextForSpecificInvalidRequestResolutions() {
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "doctor-request",
            Optional.of("--request"),
            problem(
                GridGrindProblemCode.INVALID_REQUEST, "Missing required field 'protocolVersion'"),
            new InvalidRequestException(
                "Missing required field 'protocolVersion'",
                Optional.of("protocolVersion"),
                Optional.empty(),
                Optional.empty(),
                null));

    assertEquals(
        Optional.of(
            "Add protocolVersion: \"V1\" at the request root. Rerun --doctor-request after"
                + " correcting the request."),
        failure.resolution());
  }

  @Test
  void readRequestFailureOmitsLocationsForNonReadRequestProblemContexts() {
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "execute",
            Optional.of("--request"),
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_REQUEST,
                "invalid request",
                new ProblemContext.ParseArguments(CliArgument.named("--request"))),
            new RuntimeException("invalid request"));

    assertEquals(Optional.empty(), failure.location());
  }

  @Test
  void unexpectedFailureFallsBackToTheCanonicalInternalErrorTitleWhenMessageIsBlank() {
    CliFailureReport failure =
        CliFailureReports.unexpectedFailure(
            "help", "unexpected-failure", new IllegalStateException("   "));

    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.code());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), failure.message());
    assertEquals(
        Optional.of(GridGrindProblemCode.INTERNAL_ERROR.resolution()), failure.resolution());
  }

  private static GridGrindProblemDetail.Problem problem(GridGrindProblemCode code, String message) {
    return GridGrindProblemDetail.Problem.of(
        code,
        message,
        new ProblemContext.ReadRequest(RequestInput.standardInput(), JsonLocation.unavailable()));
  }
}
