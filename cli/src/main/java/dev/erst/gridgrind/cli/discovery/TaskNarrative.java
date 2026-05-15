package dev.erst.gridgrind.cli.discovery;

import java.util.List;

/** Human-facing discovery language attached to one CLI-owned task archetype. */
public record TaskNarrative(
    String summary,
    List<String> outcomes,
    List<String> requiredInputs,
    List<String> optionalFeatures) {
  public TaskNarrative {
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    outcomes = CliDiscoveryValidation.copyStrings(outcomes, "outcomes");
    requiredInputs = CliDiscoveryValidation.copyStrings(requiredInputs, "requiredInputs");
    optionalFeatures = CliDiscoveryValidation.copyStrings(optionalFeatures, "optionalFeatures");
  }
}
