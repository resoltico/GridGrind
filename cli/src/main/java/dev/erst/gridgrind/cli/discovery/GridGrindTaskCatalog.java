package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import java.util.List;
import java.util.Optional;

/** CLI-owned task/intention catalog layered on top of the exact protocol surface. */
public final class GridGrindTaskCatalog {
  private GridGrindTaskCatalog() {}

  /** Returns the full task catalog. */
  public static TaskCatalog catalog() {
    return new TaskCatalog(
        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(), taskEntries());
  }

  /** Returns one task entry by its stable id, or empty when unknown. */
  public static Optional<TaskEntry> entryFor(String id) {
    java.util.Objects.requireNonNull(id, "id must not be null");
    String lookup = id.trim();
    if (lookup.isBlank()) {
      throw new IllegalArgumentException("id must not be blank");
    }
    return GridGrindCliRecipeRegistry.taskEntryFor(lookup);
  }

  private static List<TaskEntry> taskEntries() {
    return GridGrindCliRecipeRegistry.taskEntries();
  }

  /** Validates that every built-in task capability reference resolves against the protocol. */
  static void validateCapabilityReferences(TaskCatalog catalog) {
    for (TaskEntry task : catalog.tasks()) {
      for (TaskPhase phase : task.workflow().phases()) {
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
