package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** One staged phase inside a task descriptor, expressed in reusable protocol capabilities. */
public record TaskPhase(
    TaskPhasePurpose purpose,
    String label,
    String objective,
    List<TaskCapabilityRef> capabilityRefs,
    List<String> notes) {
  public TaskPhase {
    Objects.requireNonNull(purpose, "purpose must not be null");
    label = TextValidation.requireNonBlank(label, "label");
    objective = TextValidation.requireNonBlank(objective, "objective");
    capabilityRefs = DiscoveryValidation.copyTaskCapabilityRefs(capabilityRefs, "capabilityRefs");
    notes = DiscoveryValidation.copyStrings(notes, "notes");
    if (capabilityRefs.isEmpty()) {
      throw new IllegalArgumentException("capabilityRefs must not be empty");
    }
    for (TaskCapabilityRef capabilityRef : capabilityRefs) {
      Objects.requireNonNull(capabilityRef, "capabilityRefs must not contain nulls");
    }
  }
}
