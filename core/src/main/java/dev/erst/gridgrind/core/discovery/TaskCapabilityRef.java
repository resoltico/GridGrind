package dev.erst.gridgrind.core.discovery;

/** One reference from a task recipe phase to an existing protocol-catalog capability entry. */
public record TaskCapabilityRef(String group, String id) {
  public TaskCapabilityRef {
    group = DiscoveryValidation.requireNonBlank(group, "group");
    id = DiscoveryValidation.requireNonBlank(id, "id");
  }

  /** Returns the protocol-catalog lookup token that resolves this capability deterministically. */
  public String qualifiedId() {
    return group + ":" + id;
  }
}
