package dev.erst.gridgrind.core.semantics;

import dev.erst.gridgrind.core.metadata.Union;

/** Tagged machine-readable precondition published for one operation. */
@Union(discriminator = "type")
public sealed interface OperationPrecondition
    permits OperationPrecondition.ColumnEditsBeforeFormulaAuthoring {
  /** Requires column edits to occur before any formula is known to be present. */
  @dev.erst.gridgrind.core.metadata.Variant(
      id = "NO_COLUMN_EDITS_AFTER_FORMULA_AUTHORING",
      summary = "Require column edits before any formula has been authored.")
  record ColumnEditsBeforeFormulaAuthoring() implements OperationPrecondition {}
}
