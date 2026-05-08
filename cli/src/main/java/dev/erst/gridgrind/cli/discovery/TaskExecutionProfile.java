package dev.erst.gridgrind.cli.discovery;

import java.util.Objects;

/** Typed execution facts for one task descriptor. */
public record TaskExecutionProfile(
    TaskSourceMode sourceMode,
    TaskPersistenceMode persistenceMode,
    TaskMutationMode mutationMode,
    TaskAssetMode assetMode) {
  public TaskExecutionProfile {
    Objects.requireNonNull(sourceMode, "sourceMode must not be null");
    Objects.requireNonNull(persistenceMode, "persistenceMode must not be null");
    Objects.requireNonNull(mutationMode, "mutationMode must not be null");
    Objects.requireNonNull(assetMode, "assetMode must not be null");
  }
}
