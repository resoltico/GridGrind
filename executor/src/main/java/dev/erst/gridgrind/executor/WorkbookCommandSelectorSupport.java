package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;

/** Centralizes selector coercion and identity validation for mutation-command translation. */
final class WorkbookCommandSelectorSupport {
  private WorkbookCommandSelectorSupport() {}

  static SheetSelector.ByName sheetByName(Selector target, MutationAction action) {
    return requireTarget(target, SheetSelector.ByName.class, action.actionType());
  }

  static SheetSelector.ByNames sheetByNames(Selector target, MutationAction action) {
    return requireTarget(target, SheetSelector.ByNames.class, action.actionType());
  }

  static RangeSelector.ByRange rangeByRange(Selector target, MutationAction action) {
    return requireTarget(target, RangeSelector.ByRange.class, action.actionType());
  }

  static RowBandSelector.Span rowSpan(Selector target, MutationAction action) {
    return requireTarget(target, RowBandSelector.Span.class, action.actionType());
  }

  static RowBandSelector.Insertion rowInsertion(Selector target, MutationAction action) {
    return requireTarget(target, RowBandSelector.Insertion.class, action.actionType());
  }

  static ColumnBandSelector.Span columnSpan(Selector target, MutationAction action) {
    return requireTarget(target, ColumnBandSelector.Span.class, action.actionType());
  }

  static ColumnBandSelector.Insertion columnInsertion(Selector target, MutationAction action) {
    return requireTarget(target, ColumnBandSelector.Insertion.class, action.actionType());
  }

  static CellSelector.ByAddress cellByAddress(Selector target, MutationAction action) {
    return requireTarget(target, CellSelector.ByAddress.class, action.actionType());
  }

  static DrawingObjectSelector.ByName drawingObjectByName(Selector target, MutationAction action) {
    return requireTarget(target, DrawingObjectSelector.ByName.class, action.actionType());
  }

  static TableSelector.ByNameOnSheet tableByNameOnSheet(Selector target, MutationAction action) {
    return requireTarget(target, TableSelector.ByNameOnSheet.class, action.actionType());
  }

  static PivotTableSelector.ByNameOnSheet pivotTableByNameOnSheet(
      Selector target, MutationAction action) {
    return requireTarget(target, PivotTableSelector.ByNameOnSheet.class, action.actionType());
  }

  static NamedRangeSelector.ScopedExact namedRangeScopedExact(
      Selector target, MutationAction action) {
    return requireTarget(target, NamedRangeSelector.ScopedExact.class, action.actionType());
  }

  static void ensureTableIdentity(Selector target, StructuredMutationAction.SetTable action) {
    TableSelector.ByNameOnSheet selector = tableByNameOnSheet(target, action);
    if (!selector.name().equals(action.table().name())
        || !selector.sheetName().equals(action.table().sheetName())) {
      throw new IllegalArgumentException(
          action.actionType() + " target must match table.name and table.sheetName");
    }
  }

  static void ensurePivotTableIdentity(
      Selector target, StructuredMutationAction.SetPivotTable action) {
    PivotTableSelector.ByNameOnSheet selector = pivotTableByNameOnSheet(target, action);
    if (!selector.name().equals(action.pivotTable().name())
        || !selector.sheetName().equals(action.pivotTable().sheetName())) {
      throw new IllegalArgumentException(
          action.actionType() + " target must match pivotTable.name and pivotTable.sheetName");
    }
  }

  static void ensureNamedRangeIdentity(
      Selector target, StructuredMutationAction.SetNamedRange action) {
    NamedRangeSelector.ScopedExact selector = namedRangeScopedExact(target, action);
    if (!WorkbookCommandLayoutInputConverter.toExcelNamedRangeName(selector).equals(action.name())
        || !WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(selector)
            .equals(WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(action.scope()))) {
      throw new IllegalArgumentException(
          action.actionType() + " target must match action name and scope");
    }
  }

  private static <T extends Selector> T requireTarget(
      Selector target, Class<T> expectedType, String actionType) {
    if (expectedType.isInstance(target)) {
      return expectedType.cast(target);
    }
    throw new IllegalArgumentException(
        actionType
            + " requires target type "
            + expectedType.getSimpleName()
            + " but got "
            + target.getClass().getSimpleName());
  }
}
