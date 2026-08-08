package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Contract tests for the one plural rejected-command envelope. */
class CliDiagnosticsTest {
  @Test
  void commandErrorContainsOnlyCommandLevelFactsAndAProblemCollection() throws Exception {
    CommandError error =
        CommandErrors.invalidArguments("execute", java.util.Optional.of("--bogus"), "Unknown argument: --bogus");

    String json = new String(GridGrindCliJson.writeBytes(error), StandardCharsets.UTF_8);

    assertEquals("REJECTED", error.status());
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, error.primaryProblem().code());
    assertEquals(java.util.Optional.of("--bogus"),
        ((ProblemContext.ParseArguments) error.primaryProblem().context()).argumentName());
    assertFalse(json.contains("exitCode"));
    assertFalse(json.contains("suggestions"));
    assertFalse(json.contains("transport"));
    assertFalse(json.contains("primaryProblem"));
  }

  @Test
  void commandErrorsOrderTheirProblemsWithoutSerializingTheAllocationTieBreaker() throws Exception {
    GridGrindProblemDetail.Problem later =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "later",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.pathAtByteOffset("steps[1].stepId", 20)));
    GridGrindProblemDetail.Problem earlier =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "earlier",
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.pathAtByteOffset("steps[0].stepId", 10)));

    CommandError error = new CommandError(errorVersion(), "execute", List.of(later, earlier));
    String json = new String(GridGrindCliJson.writeBytes(error), StandardCharsets.UTF_8);

    assertEquals(List.of(earlier, later), error.problems());
    assertFalse(json.contains("ordinal"));
  }

  @Test
  void emptyProblemCollectionsAreRejected() {
    assertEquals(
        "problems must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CommandError(errorVersion(), "execute", List.of()))
            .getMessage());
  }

  private static dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion errorVersion() {
    return dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current();
  }
}
