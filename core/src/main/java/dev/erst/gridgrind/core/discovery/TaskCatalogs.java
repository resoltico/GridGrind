package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import java.util.List;
import java.util.Optional;

/** Product-owned task/intention catalog layered on top of the exact protocol surface. */
public final class TaskCatalogs {
  private TaskCatalogs() {}

  /** Returns the full task catalog. */
  public static TaskCatalog catalog() {
    return new TaskCatalog(
        dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT, taskEntries());
  }

  /** Returns one task entry by its stable id, or empty when unknown. */
  public static Optional<TaskEntry> entryFor(String id) {
    java.util.Objects.requireNonNull(id, "id must not be null");
    String lookup = id.trim();
    if (lookup.isBlank()) {
      throw new IllegalArgumentException("id must not be blank");
    }
    return RecipeRegistry.taskEntryFor(lookup);
  }

  private static List<TaskEntry> taskEntries() {
    return RecipeRegistry.taskEntries();
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
