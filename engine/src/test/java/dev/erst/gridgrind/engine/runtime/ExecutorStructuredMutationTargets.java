package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;

/**
 * Derives semantic test targets for structured mutations whose target is embedded in the action.
 */
final class ExecutorStructuredMutationTargets {
  private ExecutorStructuredMutationTargets() {}

  static ExecutorTestPlanSupport.PendingMutation namedRange(
      StructuredMutationAction.SetNamedRange action) {
    return ExecutorTestPlanSupport.mutate(
        namedRangeSelector(action.name(), action.scope()), action);
  }

  private static NamedRangeSelector namedRangeSelector(String name, NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> new NamedRangeSelector.WorkbookScope(name);
      case NamedRangeScope.Sheet sheet ->
          new NamedRangeSelector.SheetScope(name, sheet.sheetName());
    };
  }
}
