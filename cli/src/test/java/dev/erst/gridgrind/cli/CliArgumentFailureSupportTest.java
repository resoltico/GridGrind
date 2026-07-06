package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused unit coverage for CLI argument-recovery suggestions and resolutions. */
class CliArgumentFailureSupportTest {
  @Test
  void queryFailuresSuggestTaskDiscoveryCommands() {
    CliDiagnostic failure =
        CliArgumentFailureSupport.reportFor(
            new String[] {"--print-recipe-keyword-match", "--query"},
            new CliArgumentsException("--query", "Missing value for --query"));

    assertEquals(
        List.of(
            "gridgrind --print-recipe-keyword-match --query \"monthly sales dashboard\"",
            "gridgrind --print-recipe-catalog"),
        failure.suggestions());
    assertEquals(
        "Use --query only with --print-recipe-keyword-match and provide one natural-language query.",
        failure.problem().resolution());
  }

  @Test
  void mistypedFlagsOfferNearestKnownCommands() {
    CliDiagnostic failure =
        CliArgumentFailureSupport.reportFor(
            new String[] {"--versoin"},
            new CliArgumentsException("--versoin", "Unknown argument: --versoin"));

    assertTrue(
        failure.suggestions().contains("gridgrind --version"),
        "near-match guidance must include the intended flag");
    assertEquals(
        "Use one exact CLI flag. Start from --help for the synopsis, --help-protocol for the"
            + " grammar, or --help-guidance for workflow-oriented commands.",
        failure.problem().resolution());
  }

  @Test
  void distantUnknownFlagsFallBackToGeneralHelpInsteadOfGuessingWorkflowCommands() {
    CliDiagnostic failure =
        CliArgumentFailureSupport.reportFor(
            new String[] {"--bogus"},
            new CliArgumentsException("--bogus", "Unknown argument: --bogus"));

    assertEquals(
        List.of("gridgrind --help", "gridgrind --help-protocol", "gridgrind --help-guidance"),
        failure.suggestions());
  }

  @Test
  void genericArgumentFailuresFallBackToHelpCommands() {
    CliDiagnostic failure =
        CliArgumentFailureSupport.reportFor(
            new String[] {"--request", ""}, new IllegalArgumentException("bad argument shape"));

    assertEquals(
        List.of("gridgrind --help", "gridgrind --help-protocol", "gridgrind --help-guidance"),
        failure.suggestions());
    assertEquals(
        "Run gridgrind --help for the synopsis, --help-protocol for the authoritative request"
            + " contract, or --help-guidance for workflows and examples.",
        failure.problem().resolution());
  }

  @Test
  void nonUnknownArgumentFailuresUseDefaultRecoveryWithoutNearestMatchLookups() {
    CliDiagnostic failure =
        CliArgumentFailureSupport.reportFor(
            new String[] {"--license", "--license"},
            new CliArgumentsException("--license", "Duplicate argument: --license"));

    assertEquals(
        List.of("gridgrind --help", "gridgrind --help-protocol", "gridgrind --help-guidance"),
        failure.suggestions());
    assertEquals(
        "Run gridgrind --help for the synopsis, --help-protocol for the authoritative request"
            + " contract, or --help-guidance for workflows and examples.",
        failure.problem().resolution());
  }

  @Test
  void commandTemplatesStayBoundToValidInvocationFamilies() {
    assertEquals(
        List.of("gridgrind --request request.json --response response.json"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--request"));
    assertEquals(
        List.of("gridgrind --print-request-template --response request.json"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--response"));
    assertEquals(
        List.of("gridgrind --doctor-request --request request.json --response doctor.json"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--doctor-request"));
    assertEquals(
        List.of("gridgrind --print-request-template --response request.json"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--print-request-template"));
    assertEquals(
        List.of("gridgrind --print-recipe-catalog"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--print-recipe-catalog"));
    assertEquals(
        List.of("gridgrind --print-recipe --lookup WORKBOOK_HEALTH"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--print-recipe"));
    assertEquals(
        List.of("gridgrind --print-recipe-catalog"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--print-recipe-catalog"));
    assertEquals(
        List.of("gridgrind --print-recipe-keyword-match --query \"monthly sales dashboard\""),
        CliArgumentFailureSupport.commandTemplatesForFlag("--print-recipe-keyword-match"));
    assertEquals(
        List.of("gridgrind --print-protocol-catalog"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--print-protocol-catalog"));
    assertEquals(
        List.of(
            "gridgrind --print-recipe --lookup WORKBOOK_HEALTH",
            "gridgrind --print-recipe --lookup DASHBOARD",
            "gridgrind --print-recipe-catalog",
            "gridgrind --print-protocol-catalog --lookup GET_CELLS"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--lookup"));
    assertEquals(
        List.of("gridgrind --print-recipe-keyword-match --query \"monthly sales dashboard\""),
        CliArgumentFailureSupport.commandTemplatesForFlag("--query"));
    assertEquals(
        List.of("gridgrind --print-protocol-catalog"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--search"));
    assertEquals(
        List.of("gridgrind --help"), CliArgumentFailureSupport.commandTemplatesForFlag("--help"));
    assertEquals(
        List.of("gridgrind --help-protocol"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--help-protocol"));
    assertEquals(
        List.of("gridgrind --help-guidance"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--help-guidance"));
    assertEquals(
        List.of("gridgrind --version"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--version"));
    assertEquals(
        List.of("gridgrind --license"),
        CliArgumentFailureSupport.commandTemplatesForFlag("--license"));
    assertEquals(List.of(), CliArgumentFailureSupport.commandTemplatesForFlag("--bogus"));
  }
}
