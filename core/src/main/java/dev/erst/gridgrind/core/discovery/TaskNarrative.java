package dev.erst.gridgrind.core.discovery;

import java.util.List;

/** Human-facing discovery language attached to one Product-owned task archetype. */
public record TaskNarrative(
    String summary,
    List<String> outcomes,
    List<String> requiredInputs,
    List<String> optionalFeatures) {
  public TaskNarrative {
    summary = DiscoveryValidation.requireNonBlank(summary, "summary");
    outcomes = DiscoveryValidation.copyStrings(outcomes, "outcomes");
    requiredInputs = DiscoveryValidation.copyStrings(requiredInputs, "requiredInputs");
    optionalFeatures = DiscoveryValidation.copyStrings(optionalFeatures, "optionalFeatures");
  }
}
