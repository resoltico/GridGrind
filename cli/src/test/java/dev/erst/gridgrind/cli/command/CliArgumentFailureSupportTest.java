package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.result.CommandError;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests argument failures without reviving a second recovery or suggestion schema. */
class CliArgumentFailureSupportTest {
  @Test
  void namedArgumentFailureKeepsTheArgumentInTheProblemContext() {
    CommandError error =
        CommandErrors.reportFor(
            new String[] {"--bogus"},
            new CliArgumentsException("--bogus", "Unknown argument: --bogus"));

    assertEquals("REJECTED", error.status());
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, error.primaryProblem().code());
    assertEquals(
        Optional.of("--bogus"),
        ((ProblemContext.ParseArguments) error.primaryProblem().context()).argumentName());
  }

  @Test
  void genericArgumentFailureDoesNotFabricateAnArgument() {
    CommandError error =
        CommandErrors.reportFor(
            new String[] {"run", "--request", "", "--grant", "grant.json"},
            new IllegalArgumentException("bad argument shape"));

    assertEquals(
        Optional.empty(),
        ((ProblemContext.ParseArguments) error.primaryProblem().context()).argumentName());
  }
}
