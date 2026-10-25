package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.WorkbookStepTargeting;
import java.util.List;
import java.util.Optional;

/** Resolves optional target-surface metadata for machine-readable catalog type entries. */
final class TypeEntryTargetingSupport {
  private TypeEntryTargetingSupport() {}

  static List<TargetSelectorEntry> targetSelectorEntries(
      Optional<WorkbookStepTargeting.TargetSurface> targetSurface) {
    return targetSurface.stream()
        .flatMap(surface -> surface.selectorFamilies().stream())
        .map(familyInfo -> new TargetSelectorEntry(familyInfo.family(), familyInfo.typeIds()))
        .toList();
  }

  @SuppressWarnings("unchecked")
  static Optional<WorkbookStepTargeting.TargetSurface> optionalTargetSurfaceFor(
      Class<? extends Record> recordType) {
    if (MutationAction.class.isAssignableFrom(recordType)) {
      return Optional.of(
          WorkbookStepTargeting.forMutationActionType(
              (Class<? extends MutationAction>) recordType));
    }
    if (Assertion.class.isAssignableFrom(recordType)) {
      return Optional.of(
          WorkbookStepTargeting.forAssertionType((Class<? extends Assertion>) recordType));
    }
    if (InspectionQuery.class.isAssignableFrom(recordType)) {
      return Optional.of(
          WorkbookStepTargeting.forInspectionQueryType(
              (Class<? extends InspectionQuery>) recordType));
    }
    return Optional.empty();
  }
}
