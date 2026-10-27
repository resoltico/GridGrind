package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Path;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Preserves path and response-ownership boundaries independently of command parsing. */
class CliArgumentEdgeCoverageTest {
  @Test
  void responseHintNeverTargetsTheInputOrMalformedOutput() {
    assertTrue(
        CliPathArguments.responsePathHint(
                new String[] {"run", "--request", "request.json", "--response", "request.json"})
            .isEmpty());
    assertTrue(
        CliPathArguments.responsePathHint(
                new String[] {
                  "run", "--request", "a/../request.json", "--response", "request.json"
                })
            .isEmpty());
    assertTrue(
        CliPathArguments.responsePathHint(new String[] {"help", "--response", ""}).isEmpty());
    assertTrue(CliPathArguments.responsePathHint(new String[] {"help", "--response"}).isEmpty());
    assertEquals(
        Path.of("out.json"),
        CliPathArguments.responsePathHint(
                new String[] {"schema", "request", "--response", "out.json"})
            .orElseThrow());
  }

  @Test
  void stdinSentinelIsRestrictedToRequestInput() {
    assertTrue(CliPathArguments.isStandardInputPath(Optional.of(Path.of("-"))));
    assertFalse(CliPathArguments.isStandardInputPath(Optional.empty()));
    assertEquals(
        Path.of("-"), CliPathArguments.requirePathValue("--request", "-", "request", true));
    assertThrows(
        CliArgumentsException.class,
        () -> CliPathArguments.requirePathValue("--response", "-", "response", false));
  }

  @Test
  void pathValuesMustExistAsArgumentsAndContainText() {
    assertThrows(
        CliArgumentsException.class,
        () -> CliPathArguments.nextValueIndex(new String[] {"--request"}, 0, "--request"));
    assertThrows(
        CliArgumentsException.class,
        () -> CliPathArguments.requirePathValue("--request", " ", "request", true));
    assertEquals(
        1,
        CliPathArguments.nextValueIndex(
            new String[] {"--request", "request.json"}, 0, "--request"));
  }

  @Test
  void requestRootsCannotBeAmbiguous() {
    assertThrows(
        CliArgumentsException.class,
        () ->
            CliExecutionArgumentValidation.validateTerminalArguments(
                Optional.of(Path.of("-")), Optional.empty(), Optional.empty(), Optional.empty()));
    assertThrows(
        CliArgumentsException.class,
        () ->
            CliExecutionArgumentValidation.validateTerminalArguments(
                Optional.of(Path.of("r.json")),
                Optional.of(Path.of("workspace")),
                Optional.empty(),
                Optional.empty()));
    assertDoesNotThrow(
        () ->
            CliExecutionArgumentValidation.validateTerminalArguments(
                Optional.of(Path.of("-")),
                Optional.of(Path.of("workspace")),
                Optional.empty(),
                Optional.empty()));
  }

  @Test
  void prettyHintRecognizesOnlyTheSupportedFlag() {
    assertTrue(CliArguments.prettyJsonHint(new String[] {"schema", "request", "--pretty"}));
    assertFalse(CliArguments.prettyJsonHint(new String[] {"schema", "request"}));
  }
}
