package dev.erst.gridgrind.core.discovery;

import java.util.List;

/** Exact machine-readable request profile derived from one published runnable recipe. */
public record RecipeRequestProfile(
    String sourceType,
    String persistenceType,
    String executionMode,
    String journalLevel,
    String calculationStrategy,
    boolean markRecalculateOnOpen,
    int stepCount,
    List<String> mutationActionTypes,
    List<String> assertionTypes,
    List<String> inspectionQueryTypes) {
  public RecipeRequestProfile {
    sourceType = DiscoveryValidation.requireNonBlank(sourceType, "sourceType");
    persistenceType = DiscoveryValidation.requireNonBlank(persistenceType, "persistenceType");
    executionMode = DiscoveryValidation.requireNonBlank(executionMode, "executionMode");
    journalLevel = DiscoveryValidation.requireNonBlank(journalLevel, "journalLevel");
    calculationStrategy =
        DiscoveryValidation.requireNonBlank(calculationStrategy, "calculationStrategy");
    if (stepCount < 0) {
      throw new IllegalArgumentException("stepCount must not be negative");
    }
    mutationActionTypes =
        DiscoveryValidation.copyStringsAllowEmpty(mutationActionTypes, "mutationActionTypes");
    assertionTypes = DiscoveryValidation.copyStringsAllowEmpty(assertionTypes, "assertionTypes");
    inspectionQueryTypes =
        DiscoveryValidation.copyStringsAllowEmpty(inspectionQueryTypes, "inspectionQueryTypes");
  }
}
