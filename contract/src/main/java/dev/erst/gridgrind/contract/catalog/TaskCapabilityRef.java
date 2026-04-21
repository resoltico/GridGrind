package dev.erst.gridgrind.contract.catalog;

/** One reference from a task recipe phase to an existing protocol-catalog capability entry. */
public record TaskCapabilityRef(String group, String id) {
  public TaskCapabilityRef {
    group = CatalogRecordValidation.requireNonBlank(group, "group");
    id = CatalogRecordValidation.requireNonBlank(id, "id");
  }

  /** Returns the protocol-catalog lookup token that resolves this capability deterministically. */
  public String qualifiedId() {
    return group + ":" + id;
  }
}
