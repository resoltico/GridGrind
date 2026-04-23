package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Optional;

/** Internal registry that couples public task entries with richer discovery and starter plans. */
final class GridGrindTaskDefinitions {
  private static final List<TaskDefinition> DEFINITIONS = buildDefinitions();
  private static final TaskCatalog CATALOG =
      new TaskCatalog(
          dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
          DEFINITIONS.stream().map(TaskDefinition::task).toList());

  private GridGrindTaskDefinitions() {}

  static TaskCatalog catalog() {
    return CATALOG;
  }

  static List<TaskDefinition> definitions() {
    return DEFINITIONS;
  }

  static Optional<TaskDefinition> definitionFor(String id) {
    String lookup = CatalogRecordValidation.requireNonBlank(id, "id").trim();
    return DEFINITIONS.stream()
        .filter(definition -> definition.task().id().equals(lookup))
        .findFirst();
  }

  static void validateCapabilityReferences() {
    validateTaskCapabilityReferences(CATALOG.tasks());
  }

  private static List<TaskDefinition> buildDefinitions() {
    List<TaskDefinition> definitions =
        List.of(
                GridGrindReportingTaskDefinitions.definitions(),
                GridGrindWorkbookTaskDefinitions.definitions())
            .stream()
            .flatMap(List::stream)
            .toList();
    validateTaskCapabilityReferences(definitions.stream().map(TaskDefinition::task).toList());
    return definitions;
  }

  static void validateTaskCapabilityReferences(List<TaskEntry> tasks) {
    for (TaskEntry task : tasks) {
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
