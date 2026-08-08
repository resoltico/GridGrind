package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import dev.erst.gridgrind.cli.discovery.CommandError;
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
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Residual coverage for canonical CLI diagnostic branches. */
class CommandErrorsResidualTest extends GridGrindCliTestSupport {
  @Test
  void readRequestFailureUsesExecutionRepairTextForGenericInvalidRequests() {
    CommandError failure =
        CommandErrors.readRequestFailure(
            1, "execute", problem(GridGrindProblemCode.INVALID_REQUEST, "invalid request"));

    assertEquals("Fix the request data and retry the workflow.", failure.primaryProblem().resolution());
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
    CommandError failure =
        CommandErrors.readRequestFailure(
            1,
            "doctor-request",
            GridGrindProblems.fromException(
                failureCause,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.pathOnly("protocolVersion"))));

    assertEquals(
        "Add protocolVersion: \"V2\" at the request root.", failure.primaryProblem().resolution());
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
    CommandError failure =
        CommandErrors.readRequestFailure(
            1,
            "doctor-request",
            GridGrindProblems.fromException(
                failureCause,
                new ProblemContext.ReadRequest(
                    RequestInput.standardInput(), JsonLocation.pathOnly("steps[1].stepId"))));

    assertEquals(
        "Make every stepId unique. Rename or remove the duplicate value 'duplicate'.",
        failure.primaryProblem().resolution());
  }

  @Test
  void readRequestFailureLeavesNonReadRequestContextsUnwrapped() {
    CommandError failure =
        CommandErrors.readRequestFailure(
            1,
            "execute",
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_REQUEST,
                "invalid request",
                new ProblemContext.ParseArguments(CliArgument.named("--request"))));

    assertEquals(Optional.of("--request"), parseArgumentsContext(failure).argumentName());
    assertFalse(failure.primaryProblem().context() instanceof ProblemContext.ReadRequest);
  }

  @Test
  void unexpectedFailureNeverSerializesTheThrowableMessage() {
    CommandError failure =
        CommandErrors.unexpectedFailure("help", new IllegalStateException("source-secret"));

    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), failure.primaryProblem().message());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.resolution(), failure.primaryProblem().resolution());
    assertFalse(failure.primaryProblem().causes().getFirst().message().contains("source-secret"));
  }

  @Test
  void cliSourcesNoLongerContainTheLegacySheetLayoutFallbackCommand() throws IOException {
    String legacySuggestion = "gridgrind --print-protocol-catalog --search \"sheet layout\"";
    for (Path sourceFile : cliSourceFiles()) {
      assertFalse(
          Files.readString(sourceFile).contains(legacySuggestion),
          () -> "legacy canned fallback command must stay deleted from " + sourceFile);
    }
  }

  @Test
  void protocolCatalogSearchSuggestionsCoverFallbackMessageExtractionBranches() {
    GridGrindProblemDetail.Problem quotedTokenProblem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Unknown field 'zoomPercent'",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.unavailable()));
    GridGrindProblemDetail.Problem noQuotedTokenProblem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "request shape mismatch",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.unavailable()));
    GridGrindProblemDetail.Problem blankQuotedTokenProblem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Unknown type value '   '",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.unavailable()));

    assertEquals(
        Optional.of("gridgrind --print-protocol-catalog --search \"zoomPercent\""),
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(quotedTokenProblem));
    assertEquals(
        Optional.empty(),
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(noQuotedTokenProblem));
    assertEquals(
        Optional.empty(),
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(blankQuotedTokenProblem));
  }

  private static GridGrindProblemDetail.Problem problem(GridGrindProblemCode code, String message) {
    return GridGrindProblemDetail.Problem.of(
        code,
        message,
        new ProblemContext.ReadRequest(RequestInput.standardInput(), JsonLocation.unavailable()));
  }

  private static List<Path> cliSourceFiles() throws IOException {
    try (java.util.stream.Stream<Path> paths = Files.walk(locateRepoRoot().resolve("cli/src"))) {
      return paths.filter(Files::isRegularFile).toList();
    }
  }

  private static Path locateRepoRoot() {
    Path current = Path.of("").toAbsolutePath().normalize();
    while (current != null && !Files.exists(current.resolve("gradle.properties"))) {
      current = current.getParent();
    }
    if (current == null) {
      throw new AssertionError("test must run inside the GridGrind repository");
    }
    return current;
  }
}
