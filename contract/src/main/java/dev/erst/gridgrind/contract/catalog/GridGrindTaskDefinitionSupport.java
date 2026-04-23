package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Shared builders for contract-owned task definitions and starter templates. */
final class GridGrindTaskDefinitionSupport {
  private GridGrindTaskDefinitionSupport() {}

  static TaskDefinition definition(
      TaskEntry task, List<String> searchTerms, WorkbookPlan requestTemplate, List<String> notes) {
    return new TaskDefinition(task, searchTerms, requestTemplate, notes);
  }

  static WorkbookPlan rebasePlan(
      WorkbookPlan template,
      String planId,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence) {
    return new WorkbookPlan(
        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
        planId,
        source,
        persistence,
        template.execution(),
        template.formulaEnvironment(),
        template.steps());
  }

  static TaskEntry task(
      String id,
      String summary,
      List<String> intentTags,
      List<String> outcomes,
      List<String> requiredInputs,
      List<String> optionalFeatures,
      List<TaskPhase> phases,
      List<String> commonPitfalls) {
    return new TaskEntry(
        id,
        summary,
        intentTags,
        outcomes,
        requiredInputs,
        optionalFeatures,
        phases,
        commonPitfalls);
  }

  static TaskPhase phase(
      String label, String objective, List<TaskCapabilityRef> capabilityRefs, List<String> notes) {
    return new TaskPhase(label, objective, capabilityRefs, notes);
  }

  static TaskCapabilityRef ref(String group, String id) {
    return new TaskCapabilityRef(group, id);
  }

  static String protocolLookupNote(
      String exactShapeLabel, List<String> operationRefs, List<String> searchTerms) {
    return "Use "
        + joinLookups(operationRefs, "--print-protocol-catalog --operation ")
        + " for exact "
        + exactShapeLabel
        + ", or "
        + joinLookups(searchTerms, "--print-protocol-catalog --search ")
        + " when browsing nearby surfaces.";
  }

  private static String joinLookups(List<String> values, String prefix) {
    if (values.isEmpty()) {
      throw new IllegalArgumentException("values must not be empty");
    }
    List<String> lookups = values.stream().map(value -> prefix + value).toList();
    if (lookups.size() == 1) {
      return lookups.getFirst();
    }
    if (lookups.size() == 2) {
      return lookups.get(0) + " and " + lookups.get(1);
    }
    return String.join(", ", lookups.subList(0, lookups.size() - 1)) + ", and " + lookups.getLast();
  }
}
