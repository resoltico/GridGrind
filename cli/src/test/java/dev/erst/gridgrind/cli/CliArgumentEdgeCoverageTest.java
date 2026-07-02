package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import java.nio.file.Path;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct edge coverage for CLI argument helpers and protocol-catalog parser variants. */
class CliArgumentEdgeCoverageTest {
  @Test
  void protocolCatalogFullFlagIsRejectedWithScopedLookupGuidance() {
    CliArgumentsException trailingFull =
        assertThrows(
            CliArgumentsException.class,
            () -> CliArguments.parse(new String[] {"--print-protocol-catalog", "--full"}));
    assertEquals("--full", trailingFull.argument());
    assertEquals(
        "--full is no longer part of the CLI grammar; use --print-protocol-catalog --lookup"
            + " <lookup-id> for one scoped catalog payload",
        trailingFull.getMessage());
  }

  @Test
  void bareFullArgumentExplainsItsOwningPrimaryCommand() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class, () -> CliArguments.parse(new String[] {"--full"}));

    assertEquals("--full", exception.argument());
    assertEquals(
        "--full is no longer part of the CLI grammar; use --print-protocol-catalog --lookup"
            + " <lookup-id> for one scoped catalog payload",
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
  void pathHelpersParseGlobalResponseStructuredFormatAndPrettyArguments() {
    assertEquals(Optional.empty(), CliRenderArguments.outputFormat(new String[] {"help"}));
    assertEquals(
        Optional.of(CliOutputFormat.STRUCTURED),
        CliRenderArguments.outputFormat(new String[] {"--format", "structured"}));
    assertEquals(
        Optional.of(CliOutputFormat.TEXT),
        CliRenderArguments.outputFormat(new String[] {"help", "--format", "text"}));

    CliRenderArguments.GlobalResponseExtraction extraction =
        CliRenderArguments.extractGlobalResponse(
            new String[] {
              "--format", "structured", "--pretty", "--response", "report.json", "help"
            });

    assertEquals(java.util.List.of("help"), extraction.remainingArgs());
    assertEquals(Optional.of(Path.of("report.json")), extraction.responsePath());
    assertEquals(Optional.of(CliOutputFormat.STRUCTURED), extraction.outputFormat());
    assertTrue(extraction.prettyJson());
    assertEquals(1, extraction.remainingArgsArray().length);
    assertEquals("help", extraction.remainingArgsArray()[0]);

    CliArgumentsException duplicateFormat =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliRenderArguments.extractGlobalResponse(
                    new String[] {"--format", "text", "--format", "structured"}));
    assertEquals("--format", duplicateFormat.argument());
    assertEquals("Duplicate argument: --format", duplicateFormat.getMessage());

    CliArgumentsException duplicateAuthoredFormat =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliRenderArguments.outputFormat(
                    new String[] {"help", "--format", "text", "--format", "structured"}));
    assertEquals("--format", duplicateAuthoredFormat.argument());
    assertEquals("Duplicate argument: --format", duplicateAuthoredFormat.getMessage());

    CliArgumentsException invalidFormat =
        assertThrows(
            CliArgumentsException.class,
            () -> CliRenderArguments.outputFormat(new String[] {"--format", "yaml"}));
    assertEquals("--format", invalidFormat.argument());
    assertEquals("--format must be one of: text, structured", invalidFormat.getMessage());

    CliArgumentsException duplicatePretty =
        assertThrows(
            CliArgumentsException.class,
            () -> CliRenderArguments.extractGlobalResponse(new String[] {"--pretty", "--pretty"}));
    assertEquals("--pretty", duplicatePretty.argument());
    assertEquals("Duplicate argument: --pretty", duplicatePretty.getMessage());
  }

  @Test
  void formatFlagIsRejectedForJsonNativeCommands() {
    CliArgumentsException exception =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliArguments.parse(
                    new String[] {"--print-request-template", "--format", "structured"}));

    assertEquals("--format", exception.argument());
    assertEquals(
        "--format is only valid with --help, --help-protocol, --help-guidance, --version, or"
            + " --license; JSON-native commands already emit JSON and use --pretty when"
            + " indentation is desired",
        exception.getMessage());
  }

  @Test
  void renderHelpersAcceptProseFormatsAndSuppressDuplicatePrettyHints() {
    assertDoesNotThrow(() -> CliArguments.parse(new String[] {"--help", "--format", "structured"}));
    assertDoesNotThrow(() -> CliArguments.parse(new String[] {"--version", "--format", "text"}));
    assertDoesNotThrow(() -> CliArguments.parse(new String[] {"--license", "--format", "text"}));
    CliArgumentsException protocolCatalogFormat =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliRenderOptionValidation.validate(
                    new CliCommand.PrintProtocolCatalogIndex(Optional.empty()),
                    Optional.of(CliOutputFormat.TEXT)));
    assertEquals("--format", protocolCatalogFormat.argument());
    CliArgumentsException executeFormat =
        assertThrows(
            CliArgumentsException.class,
            () ->
                CliRenderOptionValidation.validate(
                    new CliCommand.Execute(
                        Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty()),
                    Optional.of(CliOutputFormat.TEXT)));
    assertEquals("--format", executeFormat.argument());
    assertTrue(CliRenderArguments.prettyJsonHint(new String[] {"--pretty"}));
    assertFalse(CliRenderArguments.prettyJsonHint(new String[] {"--pretty", "--pretty"}));
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
