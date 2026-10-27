package dev.erst.gridgrind.core.discovery;

import java.util.List;

/** Internal grouped typed discovery signals for one task descriptor. */
public record TaskInteractionProfile(
    List<TaskInputKind> requiredInputKinds, List<TaskVerificationKind> verificationKinds) {
  public TaskInteractionProfile {
    requiredInputKinds =
        DiscoveryValidation.copyEnumValues(requiredInputKinds, "requiredInputKinds", Enum::name);
    verificationKinds =
        DiscoveryValidation.copyEnumValues(verificationKinds, "verificationKinds", Enum::name);
  }
}
