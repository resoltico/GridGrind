package dev.erst.gridgrind.cli.discovery;

import java.util.Objects;

/** One high-level office-work task archetype composed from reusable protocol capabilities. */
public record TaskEntry(
    String id,
    TaskDiscoveryProfile discoveryProfile,
    TaskNarrative narrative,
    TaskExecutionProfile executionProfile,
    TaskInteractionProfile interactionProfile,
    TaskStarterContract starter,
    TaskWorkflow workflow) {
  public TaskEntry {
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    Objects.requireNonNull(discoveryProfile, "discoveryProfile must not be null");
    Objects.requireNonNull(narrative, "narrative must not be null");
    Objects.requireNonNull(executionProfile, "executionProfile must not be null");
    Objects.requireNonNull(interactionProfile, "interactionProfile must not be null");
    Objects.requireNonNull(starter, "starter must not be null");
    Objects.requireNonNull(workflow, "workflow must not be null");
  }
}
