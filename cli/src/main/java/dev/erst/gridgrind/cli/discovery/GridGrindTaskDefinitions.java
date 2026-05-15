package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import java.util.List;
import java.util.Optional;

/** Internal registry for the CLI-owned public task-descriptor catalog. */
final class GridGrindTaskDefinitions {
  private static final List<TaskEntry> ENTRIES = buildEntries();
  private static final TaskCatalog CATALOG =
      new TaskCatalog(dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(), ENTRIES);

  private GridGrindTaskDefinitions() {}

  static TaskCatalog catalog() {
    return CATALOG;
  }

  static List<TaskEntry> entries() {
    return ENTRIES;
  }

  static Optional<TaskEntry> entryFor(String id) {
    java.util.Objects.requireNonNull(id, "id must not be null");
    String lookup = id.trim();
    if (lookup.isBlank()) {
      throw new IllegalArgumentException("id must not be blank");
    }
    return ENTRIES.stream().filter(task -> task.id().equals(lookup)).findFirst();
  }

  static void validateCapabilityReferences() {
    validateTaskCapabilityReferences(CATALOG.tasks());
  }

  private static List<TaskEntry> buildEntries() {
    List<TaskEntry> entries =
        List.of(
            TabularReportTaskDefinition.entry(),
            DashboardTaskDefinition.entry(),
            DataEntryWorkflowTaskDefinition.entry(),
            PivotReportTaskDefinition.entry(),
            AuditExistingWorkbookTaskDefinition.entry(),
            CustomXmlWorkflowTaskDefinition.entry(),
            DrawingAndSignatureWorkflowTaskDefinition.entry(),
            WorkbookMaintenanceTaskDefinition.entry());
    validateTaskCapabilityReferences(entries);
    return entries;
  }

  static void validateTaskCapabilityReferences(List<TaskEntry> tasks) {
    for (TaskEntry task : tasks) {
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
