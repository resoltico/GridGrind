package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.metadata.Variant;

/** A host failure before workbook execution, identified without binding core to a transport. */
@Variant(id = "HOST_RUNTIME", summary = "A host failure before workbook execution begins.")
public record HostContext(
    @dev.erst.gridgrind.core.metadata.FieldRule(
            kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
        String hostId)
    implements ProblemContext {
  /** Validates the caller-supplied host identity. */
  public HostContext {
    java.util.Objects.requireNonNull(hostId, "hostId must not be null");
    if (hostId.isBlank()) throw new IllegalArgumentException("hostId must not be blank");
  }
}
