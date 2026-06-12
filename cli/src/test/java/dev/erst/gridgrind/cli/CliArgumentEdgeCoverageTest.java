package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import java.nio.file.Path;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct edge coverage for CLI argument helpers and protocol-catalog parser variants. */
class CliArgumentEdgeCoverageTest {
  @Test
  void protocolCatalogFullParsesIntoDedicatedCommand() {
    CliCommand.PrintProtocolCatalogAll command =
        assertInstanceOf(
            CliCommand.PrintProtocolCatalogAll.class,
            CliArguments.parse(new String[] {"--print-protocol-catalog", "--full"}));

    assertEquals(Optional.empty(), command.responsePath());
  }

  @Test
  void protocolCatalogRejectsDuplicateAndConflictingFullFlags() {
    CliArgumentsException duplicate =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(new String[] {"--print-protocol-catalog", "--full", "--full"}));
    assertEquals("--full", duplicate.argument());
    assertEquals("Duplicate argument: --full", duplicate.getMessage());

    CliArgumentsException conflictingLookup =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--print-protocol-catalog", "--full", "--lookup", "SET_CELL"}));
    assertEquals("--full", conflictingLookup.argument());
    assertEquals(
        "--print-protocol-catalog does not allow --full together with --lookup or --search",
        conflictingLookup.getMessage());

    CliArgumentsException conflictingSearch =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--print-protocol-catalog", "--search", "chart", "--full"}));
    assertEquals("--full", conflictingSearch.argument());
    assertEquals(
        "--print-protocol-catalog does not allow --full together with --lookup or --search",
        conflictingSearch.getMessage());
  }

  @Test
  void bareFullArgumentExplainsItsOwningPrimaryCommand() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class, () -> CliArguments.parse(new String[] {"--full"}));

    assertEquals("--full", exception.argument());
    assertEquals(
        "--full requires --print-protocol-catalog and emits the complete protocol catalog",
        exception.getMessage());
  }

  @Test
  void immediatePrimaryCommandsRejectTrailingExecutionArguments() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--version", "--request", "request.json"}));

    assertEquals("--request", exception.argument());
    assertEquals("--version does not allow --request", exception.getMessage());

    CliArgumentsException tempRootException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--version", "--temp-root", "scratch"}));
    assertEquals("--temp-root", tempRootException.argument());
    assertEquals("--version does not allow --temp-root", tempRootException.getMessage());

    CliArgumentsException executionRootException =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--version", "--execution-root", "workspace"}));
    assertEquals("--execution-root", executionRootException.argument());
    assertEquals("--version does not allow --execution-root", executionRootException.getMessage());
  }

  @Test
  void responsePathHintSuppressesInvalidOrSelfOverwritingTargets() {
    assertEquals(
        Optional.empty(),
        CliPathArguments.responsePathHint(
            new String[] {"--request", "request.json", "--response", "request.json"}));

    assertEquals(
        Optional.empty(),
        CliPathArguments.responsePathHint(
            new String[] {"--response", "first.json", "--response", "second.json"}));

    assertEquals(
        Optional.of(Path.of("response.json")),
        CliPathArguments.responsePathHint(
            new String[] {"--request", "-", "--response", "response.json"}));
  }

  @Test
  void pathHelpersPreserveStandardInputSentinelWhereAllowed() {
    Optional<Path> requestPath = CliPathArguments.requestPath(new String[] {"--request", "-"});

    assertEquals(Optional.of(Path.of("-")), requestPath);
    assertTrue(CliPathArguments.isStandardInputPath(requestPath));

    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () -> CliPathArguments.requirePathValue("--response", "-", "response path", false));
    assertEquals("--response", exception.argument());
    assertEquals(
        "--response does not allow the standard-input path sentinel '-'", exception.getMessage());
  }

  @Test
  void terminalArgumentValidationRejectsConflictingExecutionShapes() {
    CliArgumentsException samePath =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliExecutionArgumentValidation.validateTerminalArguments(
                    Optional.of(Path.of("request.json")),
                    Optional.empty(),
                    Optional.of(Path.of("request.json"))));
    assertEquals("--response", samePath.argument());
    assertEquals("--request and --response must not point to the same path", samePath.getMessage());

    CliArgumentsException missingExecutionRoot =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliExecutionArgumentValidation.validateTerminalArguments(
                    Optional.of(Path.of("-")), Optional.empty(), Optional.empty()));
    assertEquals("--execution-root", missingExecutionRoot.argument());
    assertEquals(
        GridGrindContractText.stdinExecutionRootRequiredMessage(),
        missingExecutionRoot.getMessage());

    CliArgumentsException fileRequestConflict =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliExecutionArgumentValidation.validateTerminalArguments(
                    Optional.of(Path.of("request.json")),
                    Optional.of(Path.of("workspace")),
                    Optional.empty()));
    assertEquals("--execution-root", fileRequestConflict.argument());
    assertEquals(
        "--execution-root cannot be combined with --request because the request file directory already owns request-root resolution",
        fileRequestConflict.getMessage());
  }
}
