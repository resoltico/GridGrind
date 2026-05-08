package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import java.util.Optional;

/** CLI-owned task/intention catalog layered on top of the exact protocol surface. */
public final class GridGrindTaskCatalog {
  private GridGrindTaskCatalog() {}

  /** Returns the full task catalog. */
  public static TaskCatalog catalog() {
    return GridGrindTaskDefinitions.catalog();
  }

  /** Returns one task entry by its stable id, or empty when unknown. */
  public static Optional<TaskEntry> entryFor(String id) {
    return GridGrindTaskDefinitions.entryFor(id);
  }

  /** Validates that every built-in task capability reference resolves against the protocol. */
  static void validateCapabilityReferences(TaskCatalog catalog) {
    for (TaskEntry task : catalog.tasks()) {
      for (TaskPhase phase : task.phases()) {
        for (TaskCapabilityRef capabilityRef : phase.capabilityRefs()) {
          if (GridGrindProtocolCatalog.entryFor(capabilityRef.qualifiedId()).isEmpty()) {
            throw new IllegalStateException(
                "Task "
                    + task.id()
                    + " references unknown protocol capability "
                    + capabilityRef.qualifiedId());
          }
        }
      }
    }
  }
}
