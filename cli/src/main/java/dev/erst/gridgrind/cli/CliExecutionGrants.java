package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.util.List;

/** Supplies the CLI's fail-closed authority when an invoker omits a host grant document. */
final class CliExecutionGrants {
  private CliExecutionGrants() {}

  static GridGrindExecutionGrant denied() {
    return new GridGrindExecutionGrant.Bounded(
        List.of(),
        List.of(),
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new GridGrindExecutionGrant.PublicationAuthority.None(),
        List.of(),
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
  }
}
