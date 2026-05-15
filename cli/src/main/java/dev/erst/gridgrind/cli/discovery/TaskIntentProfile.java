package dev.erst.gridgrind.cli.discovery;

import java.util.List;

/** Typed task intent facts that agents can reason about without parsing prose. */
public record TaskIntentProfile(List<TaskGoalKind> goals, List<TaskArtifactKind> artifacts) {
  public TaskIntentProfile {
    goals = CliDiscoveryValidation.copyEnumValues(goals, "goals", Enum::name);
    artifacts = CliDiscoveryValidation.copyEnumValues(artifacts, "artifacts", Enum::name);
    if (goals.isEmpty()) {
      throw new IllegalArgumentException("goals must not be empty");
    }
    if (artifacts.isEmpty()) {
      throw new IllegalArgumentException("artifacts must not be empty");
    }
  }
}
