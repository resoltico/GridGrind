package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.Assertion;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.MutationAction;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import java.util.Objects;
import java.util.Optional;

/** Mutation and assertion diagnostic field extraction for execution journal contexts. */
final class ExecutionActionDiagnosticFields {
  private ExecutionActionDiagnosticFields() {}

  static Optional<String> sheetNameFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof StructuredMutationAction.SetTable setTable) {
      return Optional.of(setTable.table().sheetName().value());
    }
    if (action instanceof StructuredMutationAction.SetPivotTable setPivotTable) {
      return Optional.of(setPivotTable.pivotTable().sheetName().value());
    }
    if (action instanceof StructuredMutationAction.SetNamedRange setNamedRange) {
      return switch (setNamedRange.target()) {
        case dev.erst.gridgrind.core.model.formula.NamedRangeTarget.Range range ->
            Optional.of(range.sheetName().value());
        case dev.erst.gridgrind.core.model.formula.NamedRangeTarget.Formula _ ->
            switch (setNamedRange.scope()) {
              case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook _ ->
                  Optional.empty();
              case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet sheet ->
                  Optional.of(sheet.sheetName().value());
            };
      };
    }
    return Optional.empty();
  }

  static Optional<String> addressFor(MutationAction action) {
    if (action instanceof StructuredMutationAction.SetPivotTable setPivotTable) {
      return Optional.of(setPivotTable.pivotTable().anchor().topLeftAddress().value());
    }
    return Optional.empty();
  }

  static Optional<String> rangeFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof StructuredMutationAction.SetTable setTable) {
      return Optional.of(setTable.table().range().value());
    }
    if (action instanceof StructuredMutationAction.SetPivotTable setPivotTable) {
      if (setPivotTable.pivotTable().source()
          instanceof PivotTableSpec.PivotSourceSpec.Range range) {
        return Optional.of(range.range().value());
      }
      return Optional.empty();
    }
    if (action instanceof StructuredMutationAction.SetNamedRange setNamedRange) {
      return switch (setNamedRange.target()) {
        case dev.erst.gridgrind.core.model.formula.NamedRangeTarget.Range range ->
            Optional.of(range.range().value());
        case dev.erst.gridgrind.core.model.formula.NamedRangeTarget.Formula _ -> Optional.empty();
      };
    }
    return Optional.empty();
  }

  static Optional<String> formulaFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof CellMutationAction.SetCell setCell) {
      return setCell.value() instanceof CellValueSpec.Formula formula
          ? inlineFormula(formula)
          : Optional.empty();
    }
    if (action instanceof StructuredMutationAction.SetNamedRange setNamedRange) {
      return switch (setNamedRange.target()) {
        case dev.erst.gridgrind.core.model.formula.NamedRangeTarget.Formula formula ->
            Optional.of(formula.formula());
        case dev.erst.gridgrind.core.model.formula.NamedRangeTarget.Range _ -> Optional.empty();
      };
    }
    return Optional.empty();
  }

  static Optional<String> formulaFor(Assertion assertion) {
    if (assertion instanceof CellAssertion.FormulaText formulaText) {
      return Optional.of(formulaText.formula());
    }
    return Optional.empty();
  }

  static Optional<String> namedRangeNameFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof StructuredMutationAction.SetNamedRange setNamedRange) {
      return Optional.of(setNamedRange.name().value());
    }
    if (action instanceof StructuredMutationAction.SetPivotTable setPivotTable
        && setPivotTable.pivotTable().source()
            instanceof PivotTableSpec.PivotSourceSpec.NamedRange namedRange) {
      return Optional.of(namedRange.name().value());
    }
    return Optional.empty();
  }

  private static Optional<String> inlineFormula(CellValueSpec.Formula formula) {
    return switch (formula.value()) {
      case dev.erst.gridgrind.core.source.TextSourceSpec.Inline inline ->
          Optional.of(inline.text());
      case dev.erst.gridgrind.core.source.TextSourceSpec.Utf8File _ -> Optional.empty();
      case dev.erst.gridgrind.core.source.TextSourceSpec.Stdin _ -> Optional.empty();
    };
  }
}
