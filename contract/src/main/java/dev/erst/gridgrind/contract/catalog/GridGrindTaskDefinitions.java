package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Optional;

/** Internal registry that couples public task entries with richer discovery and starter plans. */
final class GridGrindTaskDefinitions {
  private static final String PROTOCOL_OPERATION_PREFIX = "--print-protocol-catalog --operation ";
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
    validateStarterTemplates(definitions);
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

  static void validateStarterTemplates(List<TaskDefinition> definitions) {
    for (TaskDefinition definition : definitions) {
      validateStarterTemplatePlaceholders(definition);
      validateStarterTemplateLookups(definition);
    }
  }

  static void validateStarterTemplatePlaceholders(TaskDefinition definition) {
    String requestText = definition.requestTemplate().toString();
    if (requestText.contains("TODO_")) {
      throw new IllegalStateException(
          "Task "
              + definition.task().id()
              + " starter request template contains unresolved TODO placeholders");
    }
    for (String note : definition.authoringNotes()) {
      if (note.contains("TODO_")) {
        throw new IllegalStateException(
            "Task "
                + definition.task().id()
                + " authoring note contains unresolved TODO placeholders");
      }
    }
  }

  static void validateStarterTemplateLookups(TaskDefinition definition) {
    for (String note : definition.authoringNotes()) {
      int searchIndex = 0;
      while (true) {
        int prefixIndex = note.indexOf(PROTOCOL_OPERATION_PREFIX, searchIndex);
        if (prefixIndex < 0) {
          break;
        }
        int valueStart = prefixIndex + PROTOCOL_OPERATION_PREFIX.length();
        int valueEnd = valueStart;
        while (valueEnd < note.length() && !Character.isWhitespace(note.charAt(valueEnd))) {
          valueEnd++;
        }
        String lookupId = note.substring(valueStart, valueEnd).replaceAll("[.,)]+$", "");
        if (GridGrindProtocolCatalog.lookupValueFor(lookupId).isEmpty()) {
          throw new IllegalStateException(
              "Task "
                  + definition.task().id()
                  + " authoring note references unknown protocol lookup "
                  + lookupId);
        }
        searchIndex = valueEnd;
      }
    }
  }
}
