package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.nio.file.Path;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for execution-parser unknown-argument handling. */
class CliExecutionCommandParserTest {
  @Test
  void unknownArgumentExceptionRejectsRemovedTaskPlanFlagAsUnknown() {
    CliArgumentsException failure =
        CliExecutionCommandParser.unknownArgumentException("--print-task-plan");

    assertEquals("--print-task-plan", failure.argument());
    assertEquals("Unknown argument: --print-task-plan", failure.getMessage());
  }

  @Test
  void rejectsConflictingRequestRootAndDuplicateGrantArguments() {
    CliArgumentsException rootFailure =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliExecutionArgumentValidation.validateTerminalArguments(
                    Optional.of(Path.of("request.json")),
                    Optional.of(Path.of("workspace")),
                    Optional.of(Path.of("grant.json")),
                    Optional.empty()));

    assertEquals("--execution-root", rootFailure.argument());
  }

  @Test
  void rejectsDuplicateHostAuthorityArgumentsBeforeExecution() {
    CliArgumentsException duplicateGrant =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliExecutionCommandParser.parse(
                    new String[] {"--grant", "first.json", "--grant", "second.json"},
                    Optional.empty()));
    CliArgumentsException duplicateSecretsProvider =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliExecutionCommandParser.parse(
                    new String[] {"--secrets-provider", "first", "--secrets-provider", "second"},
                    Optional.empty()));

    assertEquals("--grant", duplicateGrant.argument());
    assertEquals("--secrets-provider", duplicateSecretsProvider.argument());
  }
}
