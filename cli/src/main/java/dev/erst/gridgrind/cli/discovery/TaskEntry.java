package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** One high-level office-work task archetype composed from reusable protocol capabilities. */
public record TaskEntry(
    String id,
    String summary,
    TaskExecutionProfile executionProfile,
    List<TaskInputKind> requiredInputKinds,
    List<TaskVerificationKind> verificationKinds,
    List<String> intentTags,
    List<String> outcomes,
    List<String> requiredInputs,
    List<String> optionalFeatures,
    List<TaskPhase> phases,
    List<String> commonPitfalls) {
  public TaskEntry {
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(executionProfile, "executionProfile must not be null");
    requiredInputKinds =
        CliDiscoveryValidation.copyEnumValues(requiredInputKinds, "requiredInputKinds", Enum::name);
    verificationKinds =
        CliDiscoveryValidation.copyEnumValues(verificationKinds, "verificationKinds", Enum::name);
    intentTags = CliDiscoveryValidation.copyStrings(intentTags, "intentTags");
    outcomes = CliDiscoveryValidation.copyStrings(outcomes, "outcomes");
    requiredInputs = CliDiscoveryValidation.copyStrings(requiredInputs, "requiredInputs");
    optionalFeatures = CliDiscoveryValidation.copyStrings(optionalFeatures, "optionalFeatures");
    phases = CliDiscoveryValidation.copyTaskPhases(phases, "phases");
    commonPitfalls = CliDiscoveryValidation.copyStrings(commonPitfalls, "commonPitfalls");
    if (phases.isEmpty()) {
      throw new IllegalArgumentException("phases must not be empty");
    }
    for (TaskPhase phase : phases) {
      Objects.requireNonNull(phase, "phases must not contain nulls");
    }
  }
}
