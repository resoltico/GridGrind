package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused unit coverage for CLI argument-recovery suggestions and resolutions. */
class CliArgumentFailureSupportTest {
  @Test
  void queryFailuresSuggestTaskDiscoveryCommands() {
    CliFailureReport failure =
        CliArgumentFailureSupport.reportFor(
            new CliArgumentsException("--query", "Missing value for --query"));

    assertEquals(
        List.of(
            "gridgrind --print-task-keyword-match --query \"monthly sales dashboard\"",
            "gridgrind --print-task-catalog"),
        failure.suggestions());
    assertEquals(
        Optional.of(
            "Use --query only with --print-task-keyword-match and provide one natural-language query."),
        failure.resolution());
  }

  @Test
  void mistypedFlagsOfferNearestKnownCommands() {
    CliFailureReport failure =
        CliArgumentFailureSupport.reportFor(
            new CliArgumentsException("--versoin", "Unknown argument: --versoin"));

    assertTrue(
        failure.suggestions().contains("gridgrind --version"),
        "near-match guidance must include the intended flag");
    assertEquals(
        Optional.of(
            "Use one exact CLI flag. Start from --help for the synopsis, --help-protocol for the"
                + " grammar, or --help-guidance for workflow-oriented commands."),
        failure.resolution());
  }

  @Test
  void genericArgumentFailuresFallBackToHelpCommands() {
    CliFailureReport failure =
        CliArgumentFailureSupport.reportFor(new IllegalArgumentException("bad argument shape"));

    assertEquals(
        List.of("gridgrind --help", "gridgrind --help-protocol", "gridgrind --help-guidance"),
        failure.suggestions());
    assertEquals(
        Optional.of(
            "Run gridgrind --help for the synopsis, --help-protocol for the authoritative request"
                + " contract, or --help-guidance for workflows and examples."),
        failure.resolution());
  }

  @Test
  void nonUnknownArgumentFailuresUseDefaultRecoveryWithoutNearestMatchLookups() {
    CliFailureReport failure =
        CliArgumentFailureSupport.reportFor(
            new CliArgumentsException("--license", "Duplicate argument: --license"));

    assertEquals(
        List.of("gridgrind --help", "gridgrind --help-protocol", "gridgrind --help-guidance"),
        failure.suggestions());
    assertEquals(
        Optional.of(
            "Run gridgrind --help for the synopsis, --help-protocol for the authoritative request"
                + " contract, or --help-guidance for workflows and examples."),
        failure.resolution());
  }
}
