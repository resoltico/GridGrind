package dev.erst.gridgrind.cli.discovery;

/** One reference from a task recipe phase to an existing protocol-catalog capability entry. */
public record TaskCapabilityRef(String group, String id) {
  public TaskCapabilityRef {
    group = CliDiscoveryValidation.requireNonBlank(group, "group");
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
  }

  /** Returns the protocol-catalog lookup token that resolves this capability deterministically. */
  public String qualifiedId() {
    return group + ":" + id;
  }
}
