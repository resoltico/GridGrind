package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** One reference from a task recipe phase to an existing protocol-catalog capability entry. */
public record TaskCapabilityRef(String group, String id) {
  public TaskCapabilityRef {
    group = TextValidation.requireNonBlank(group, "group");
    id = TextValidation.requireNonBlank(id, "id");
  }

  /** Returns the protocol-catalog lookup token that resolves this capability deterministically. */
  public String qualifiedId() {
    return group + ":" + id;
  }
}
