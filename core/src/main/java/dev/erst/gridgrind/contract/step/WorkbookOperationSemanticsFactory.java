package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.catalog.OperationEffect;
import dev.erst.gridgrind.contract.catalog.OperationEffectFootprint;
import dev.erst.gridgrind.contract.catalog.OperationPrecondition;
import dev.erst.gridgrind.contract.catalog.OperationSemantics;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.List;

/** Derives the canonical static effects and preconditions for one concrete operation record. */
final class WorkbookOperationSemanticsFactory {
  private WorkbookOperationSemanticsFactory() {}

  static OperationSemantics semanticsFor(Class<? extends Record> operationType) {
    if (CellMutationAction.class.isAssignableFrom(operationType)) {
      return mutation(
          operationType,
          List.of(
              OperationEffect.CELLS,
              OperationEffect.FORMULAS,
              OperationEffect.STYLES,
              OperationEffect.HYPERLINKS,
              OperationEffect.COMMENTS,
              OperationEffect.DRAWINGS));
    }
    if (DrawingMutationAction.class.isAssignableFrom(operationType)) {
      return mutation(
          operationType,
          List.of(OperationEffect.DRAWINGS, OperationEffect.CELLS, OperationEffect.FORMULAS));
    }
    if (StructuredMutationAction.class.isAssignableFrom(operationType)) {
      return mutation(
          operationType,
          List.of(
              OperationEffect.TABLES,
              OperationEffect.NAMED_RANGES,
              OperationEffect.DATA_VALIDATIONS,
              OperationEffect.CONDITIONAL_FORMATTING,
              OperationEffect.AUTOFILTERS,
              OperationEffect.PIVOT_TABLES,
              OperationEffect.CUSTOM_XML,
              OperationEffect.FORMULAS,
              OperationEffect.CELLS));
    }
    if (WorkbookMutationAction.class.isAssignableFrom(operationType)) {
      return mutation(
          operationType,
          List.of(
              OperationEffect.WORKBOOK_STRUCTURE,
              OperationEffect.SHEET_STRUCTURE,
              OperationEffect.ROWS,
              OperationEffect.COLUMNS,
              OperationEffect.CELLS,
              OperationEffect.FORMULAS,
              OperationEffect.STYLES,
              OperationEffect.HYPERLINKS,
              OperationEffect.COMMENTS,
              OperationEffect.DRAWINGS,
              OperationEffect.TABLES,
              OperationEffect.NAMED_RANGES,
              OperationEffect.DATA_VALIDATIONS,
              OperationEffect.CONDITIONAL_FORMATTING,
              OperationEffect.AUTOFILTERS,
              OperationEffect.PIVOT_TABLES,
              OperationEffect.WORKBOOK_METADATA,
              OperationEffect.WORKBOOK_SECURITY));
    }
    if (Assertion.class.isAssignableFrom(operationType)
        || InspectionQuery.class.isAssignableFrom(operationType)) {
      return new OperationSemantics(
          List.of(OperationEffect.READ_WORKBOOK),
          OperationEffectFootprint.TARGET_BOUNDED,
          List.of());
    }
    throw new IllegalArgumentException("operation type has no static semantics: " + operationType);
  }

  private static OperationSemantics mutation(
      Class<? extends Record> operationType, List<OperationEffect> effects) {
    List<OperationEffect> combined = new java.util.ArrayList<>(effects.size() + 2);
    combined.add(OperationEffect.READ_WORKBOOK);
    combined.add(OperationEffect.MUTATE_WORKBOOK);
    combined.addAll(effects);
    return new OperationSemantics(
        combined,
        OperationEffectFootprint.CONSERVATIVE_WORKBOOK_WIDE,
        staticPreconditions(operationType));
  }

  private static List<OperationPrecondition> staticPreconditions(
      Class<? extends Record> operationType) {
    if (operationType == WorkbookMutationAction.InsertColumns.class
        || operationType == WorkbookMutationAction.DeleteColumns.class
        || operationType == WorkbookMutationAction.ShiftColumns.class) {
      return List.of(new OperationPrecondition.ColumnEditsBeforeFormulaAuthoring());
    }
    return List.of();
  }
}
