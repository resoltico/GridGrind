package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;
import java.util.Optional;

/** Mutation and assertion diagnostic field extraction for execution journal contexts. */
final class ExecutionActionDiagnosticFields {
  private ExecutionActionDiagnosticFields() {}

  static Optional<String> sheetNameFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof StructuredMutationAction.SetTable setTable) {
      return Optional.of(setTable.table().sheetName());
    }
    if (action instanceof StructuredMutationAction.SetPivotTable setPivotTable) {
      return Optional.of(setPivotTable.pivotTable().sheetName());
    }
    if (action instanceof StructuredMutationAction.SetNamedRange setNamedRange) {
      return Optional.ofNullable(setNamedRange.target().sheetName())
          .or(() -> sheetNameFor(setNamedRange.scope()));
    }
    return Optional.empty();
  }

  static Optional<String> addressFor(MutationAction action) {
    if (action instanceof StructuredMutationAction.SetPivotTable setPivotTable) {
      return Optional.of(setPivotTable.pivotTable().anchor().topLeftAddress());
    }
    return Optional.empty();
  }

  static Optional<String> rangeFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof StructuredMutationAction.SetTable setTable) {
      return Optional.of(setTable.table().range());
    }
    if (action instanceof StructuredMutationAction.SetPivotTable setPivotTable) {
      if (setPivotTable.pivotTable().source() instanceof PivotTableInput.Source.Range range) {
        return Optional.of(range.range());
      }
      return Optional.empty();
    }
    if (action instanceof StructuredMutationAction.SetNamedRange setNamedRange) {
      return Optional.ofNullable(setNamedRange.target().range());
    }
    if (action
        instanceof StructuredMutationAction.SetConditionalFormatting setConditionalFormatting) {
      return setConditionalFormatting.conditionalFormatting().ranges().size() == 1
          ? Optional.of(setConditionalFormatting.conditionalFormatting().ranges().getFirst())
          : Optional.empty();
    }
    return Optional.empty();
  }

  static Optional<String> formulaFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof CellMutationAction.SetCell setCell) {
      return setCell.value() instanceof CellInput.Formula formula
          ? inlineFormula(formula)
          : Optional.empty();
    }
    if (action instanceof StructuredMutationAction.SetNamedRange setNamedRange) {
      return Optional.ofNullable(setNamedRange.target().formula());
    }
    return Optional.empty();
  }

  static Optional<String> formulaFor(Assertion assertion) {
    if (assertion instanceof Assertion.FormulaText formulaText) {
      return Optional.of(formulaText.formula());
    }
    return Optional.empty();
  }

  static Optional<String> namedRangeNameFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof StructuredMutationAction.SetNamedRange setNamedRange) {
      return Optional.of(setNamedRange.name());
    }
    if (action instanceof StructuredMutationAction.SetPivotTable setPivotTable
        && setPivotTable.pivotTable().source()
            instanceof PivotTableInput.Source.NamedRange namedRange) {
      return Optional.of(namedRange.name());
    }
    return Optional.empty();
  }

  private static Optional<String> inlineFormula(CellInput.Formula formula) {
    if (formula.source() instanceof TextSourceInput.Inline inline) {
      return Optional.of(inline.text());
    }
    return Optional.empty();
  }

  private static Optional<String> sheetNameFor(NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> Optional.empty();
      case NamedRangeScope.Sheet sheet -> Optional.of(sheet.sheetName());
    };
  }
}
