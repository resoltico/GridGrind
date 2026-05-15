package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** Machine-facing discovery profile for one CLI-owned task descriptor. */
public record TaskDiscoveryProfile(
    List<String> discoveryTerms, List<String> intentTags, TaskIntentProfile intentProfile) {
  public TaskDiscoveryProfile {
    discoveryTerms = CliDiscoveryValidation.copyStrings(discoveryTerms, "discoveryTerms");
    intentTags = CliDiscoveryValidation.copyStrings(intentTags, "intentTags");
    Objects.requireNonNull(intentProfile, "intentProfile must not be null");
    if (discoveryTerms.isEmpty()) {
      throw new IllegalArgumentException("discoveryTerms must not be empty");
    }
  }
}
