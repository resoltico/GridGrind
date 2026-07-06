package dev.erst.gridgrind.cli.discovery;

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
    sourceType = CliDiscoveryValidation.requireNonBlank(sourceType, "sourceType");
    persistenceType = CliDiscoveryValidation.requireNonBlank(persistenceType, "persistenceType");
    executionMode = CliDiscoveryValidation.requireNonBlank(executionMode, "executionMode");
    journalLevel = CliDiscoveryValidation.requireNonBlank(journalLevel, "journalLevel");
    calculationStrategy =
        CliDiscoveryValidation.requireNonBlank(calculationStrategy, "calculationStrategy");
    if (stepCount < 0) {
      throw new IllegalArgumentException("stepCount must not be negative");
    }
    mutationActionTypes =
        CliDiscoveryValidation.copyStringsAllowEmpty(mutationActionTypes, "mutationActionTypes");
    assertionTypes = CliDiscoveryValidation.copyStringsAllowEmpty(assertionTypes, "assertionTypes");
    inspectionQueryTypes =
        CliDiscoveryValidation.copyStringsAllowEmpty(inspectionQueryTypes, "inspectionQueryTypes");
  }
}
