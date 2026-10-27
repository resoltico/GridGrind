package dev.erst.gridgrind.core.discovery;

import java.util.List;
import java.util.Objects;

/** Machine-facing discovery profile for one Product-owned task descriptor. */
public record TaskDiscoveryProfile(List<String> discoveryTerms, TaskIntentProfile intentProfile) {
  public TaskDiscoveryProfile {
    discoveryTerms = DiscoveryValidation.copyStrings(discoveryTerms, "discoveryTerms");
    Objects.requireNonNull(intentProfile, "intentProfile must not be null");
    if (discoveryTerms.isEmpty()) {
      throw new IllegalArgumentException("discoveryTerms must not be empty");
    }
  }
}
