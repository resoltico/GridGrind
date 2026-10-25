package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Tagged machine-readable precondition published for one operation. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes.Type(
    value = OperationPrecondition.ColumnEditsBeforeFormulaAuthoring.class,
    name = "NO_COLUMN_EDITS_AFTER_FORMULA_AUTHORING")
public sealed interface OperationPrecondition
    permits OperationPrecondition.ColumnEditsBeforeFormulaAuthoring {
  /** Requires column edits to occur before any formula is known to be present. */
  record ColumnEditsBeforeFormulaAuthoring() implements OperationPrecondition {}
}
