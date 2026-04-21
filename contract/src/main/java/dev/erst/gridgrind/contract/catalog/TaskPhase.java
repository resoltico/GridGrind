package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

/** One staged phase inside a task descriptor, expressed in reusable protocol capabilities. */
public record TaskPhase(
    String label, String objective, List<TaskCapabilityRef> capabilityRefs, List<String> notes) {
  public TaskPhase {
    label = CatalogRecordValidation.requireNonBlank(label, "label");
    objective = CatalogRecordValidation.requireNonBlank(objective, "objective");
    capabilityRefs =
        CatalogRecordValidation.copyTaskCapabilityRefs(capabilityRefs, "capabilityRefs");
    notes = CatalogRecordValidation.copyStrings(notes, "notes");
    if (capabilityRefs.isEmpty()) {
      throw new IllegalArgumentException("capabilityRefs must not be empty");
    }
    for (TaskCapabilityRef capabilityRef : capabilityRefs) {
      Objects.requireNonNull(capabilityRef, "capabilityRefs must not contain nulls");
    }
  }
}
