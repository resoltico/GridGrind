package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** One high-level office-work task archetype composed from reusable protocol capabilities. */
public record TaskEntry(
    String id,
    List<String> intentTags,
    TaskDiscoveryProfile discoveryProfile,
    TaskNarrative narrative,
    TaskExecutionProfile executionProfile,
    TaskInteractionProfile interactionProfile,
    TaskStarterContract starter,
    TaskWorkflow workflow) {
  public TaskEntry {
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    intentTags = CliDiscoveryValidation.copyStrings(intentTags, "intentTags");
    Objects.requireNonNull(discoveryProfile, "discoveryProfile must not be null");
    Objects.requireNonNull(narrative, "narrative must not be null");
    Objects.requireNonNull(executionProfile, "executionProfile must not be null");
    Objects.requireNonNull(interactionProfile, "interactionProfile must not be null");
    Objects.requireNonNull(starter, "starter must not be null");
    Objects.requireNonNull(workflow, "workflow must not be null");
  }
}
