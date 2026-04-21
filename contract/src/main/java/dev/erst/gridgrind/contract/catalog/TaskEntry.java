package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

/** One high-level office-work task archetype composed from reusable protocol capabilities. */
public record TaskEntry(
    String id,
    String summary,
    List<String> intentTags,
    List<String> outcomes,
    List<String> requiredInputs,
    List<String> optionalFeatures,
    List<TaskPhase> phases,
    List<String> commonPitfalls) {
  public TaskEntry {
    id = CatalogRecordValidation.requireNonBlank(id, "id");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    intentTags = CatalogRecordValidation.copyStrings(intentTags, "intentTags");
    outcomes = CatalogRecordValidation.copyStrings(outcomes, "outcomes");
    requiredInputs = CatalogRecordValidation.copyStrings(requiredInputs, "requiredInputs");
    optionalFeatures = CatalogRecordValidation.copyStrings(optionalFeatures, "optionalFeatures");
    phases = CatalogRecordValidation.copyTaskPhases(phases, "phases");
    commonPitfalls = CatalogRecordValidation.copyStrings(commonPitfalls, "commonPitfalls");
    if (phases.isEmpty()) {
      throw new IllegalArgumentException("phases must not be empty");
    }
    for (TaskPhase phase : phases) {
      Objects.requireNonNull(phase, "phases must not contain nulls");
    }
  }
}
