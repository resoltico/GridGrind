package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
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
    sourceType = TextValidation.requireNonBlank(sourceType, "sourceType");
    persistenceType = TextValidation.requireNonBlank(persistenceType, "persistenceType");
    executionMode = TextValidation.requireNonBlank(executionMode, "executionMode");
    journalLevel = TextValidation.requireNonBlank(journalLevel, "journalLevel");
    calculationStrategy =
        TextValidation.requireNonBlank(calculationStrategy, "calculationStrategy");
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
