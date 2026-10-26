package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;

/**
 * Derives semantic test targets for structured mutations whose target is embedded in the action.
 */
final class ExecutorStructuredMutationTargets {
  private ExecutorStructuredMutationTargets() {}

  static ExecutorTestPlanSupport.PendingMutation namedRange(
      StructuredMutationAction.SetNamedRange action) {
    return ExecutorTestPlanSupport.mutate(
        namedRangeSelector(action.name().value(), action.scope()), action);
  }

  private static NamedRangeSelector namedRangeSelector(String name, NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ ->
          new NamedRangeSelector.WorkbookScope(new DefinedName(name));
      case NamedRangeScope.Sheet sheet ->
          new NamedRangeSelector.SheetScope(new DefinedName(name), sheet.sheetName());
    };
  }
}
