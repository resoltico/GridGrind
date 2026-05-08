package dev.erst.gridgrind.cli.discovery;

import java.util.List;

/** Internal grouped typed discovery signals for one task descriptor. */
record TaskInteractionProfile(
    List<TaskInputKind> requiredInputKinds, List<TaskVerificationKind> verificationKinds) {
  TaskInteractionProfile {
    requiredInputKinds =
        CliDiscoveryValidation.copyEnumValues(requiredInputKinds, "requiredInputKinds", Enum::name);
    verificationKinds =
        CliDiscoveryValidation.copyEnumValues(verificationKinds, "verificationKinds", Enum::name);
  }
}
