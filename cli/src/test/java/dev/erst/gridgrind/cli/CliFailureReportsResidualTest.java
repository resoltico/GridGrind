package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.json.DuplicateStepId;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.MissingRequiredField;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
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
            "doctor-request",
            Optional.of("--request"),
            GridGrindProblems.fromException(
                failureCause,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.pathOnly("protocolVersion"))),
            failureCause);

    assertEquals(
        Optional.of(
            "Add protocolVersion: \"V1\" at the request root. Use"
                + " --print-protocol-catalog --search \"sheet layout\" or --help-protocol when"
                + " you need the authoritative field and discriminator contract."),
        failure.resolution());
  }

  @Test
  void readRequestFailureUsesDoctorRepairTextForSpecificInvalidRequestInvariants() {
    InvalidRequestException failureCause =
        new InvalidRequestException(
            new DuplicateStepId("duplicate", "steps[1].stepId"),
            Optional.of("steps[1].stepId"),
            Optional.empty(),
            Optional.empty(),
            null);
    CliFailureReport failure =
        CliFailureReports.readRequestFailure(
            1,
            "doctor-request",
            Optional.of("--request"),
            GridGrindProblems.fromException(
                failureCause,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.pathOnly("steps[1].stepId"))),
            failureCause);

    assertEquals(
        Optional.of(
            "Make every stepId unique. Rename or remove the duplicate value 'duplicate'."
                + " Rerun --doctor-request after correcting the request."),
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
