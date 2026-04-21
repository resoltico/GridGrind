package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

/** Mutation and assertion diagnostic field extraction for execution journal contexts. */
final class ExecutionActionDiagnosticFields {
  private ExecutionActionDiagnosticFields() {}

  static String sheetNameFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof MutationAction.SetTable setTable) {
      return setTable.table().sheetName();
    }
    if (action instanceof MutationAction.SetPivotTable setPivotTable) {
      return setPivotTable.pivotTable().sheetName();
    }
    if (action instanceof MutationAction.SetNamedRange setNamedRange) {
      return setNamedRange.target().sheetName() != null
          ? setNamedRange.target().sheetName()
          : sheetNameFor(setNamedRange.scope());
    }
    return null;
  }

  static String addressFor(MutationAction action) {
    if (action instanceof MutationAction.SetPivotTable setPivotTable) {
      return setPivotTable.pivotTable().anchor().topLeftAddress();
    }
    return null;
  }

  static String rangeFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof MutationAction.SetTable setTable) {
      return setTable.table().range();
    }
    if (action instanceof MutationAction.SetPivotTable setPivotTable) {
      if (setPivotTable.pivotTable().source() instanceof PivotTableInput.Source.Range range) {
        return range.range();
      }
      return null;
    }
    if (action instanceof MutationAction.SetNamedRange setNamedRange) {
      return setNamedRange.target().range();
    }
    if (action instanceof MutationAction.SetConditionalFormatting setConditionalFormatting) {
      return setConditionalFormatting.conditionalFormatting().ranges().size() == 1
          ? setConditionalFormatting.conditionalFormatting().ranges().getFirst()
          : null;
    }
    return null;
  }

  static String formulaFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof MutationAction.SetCell setCell) {
      return setCell.value() instanceof CellInput.Formula formula ? inlineFormula(formula) : null;
    }
    if (action instanceof MutationAction.SetNamedRange setNamedRange) {
      return setNamedRange.target().formula();
    }
    return null;
  }

  static String formulaFor(Assertion assertion) {
    if (assertion instanceof Assertion.FormulaText formulaText) {
      return formulaText.formula();
    }
    return null;
  }

  static String namedRangeNameFor(MutationAction action) {
    Objects.requireNonNull(action, "action must not be null");
    if (action instanceof MutationAction.SetNamedRange setNamedRange) {
      return setNamedRange.name();
    }
    if (action instanceof MutationAction.SetPivotTable setPivotTable
        && setPivotTable.pivotTable().source()
            instanceof PivotTableInput.Source.NamedRange namedRange) {
      return namedRange.name();
    }
    return null;
  }

  private static String inlineFormula(CellInput.Formula formula) {
    if (formula.source() instanceof TextSourceInput.Inline inline) {
      return inline.text();
    }
    return null;
  }

  private static String sheetNameFor(NamedRangeScope scope) {
    return switch (scope) {
      case NamedRangeScope.Workbook _ -> null;
      case NamedRangeScope.Sheet sheet -> sheet.sheetName();
    };
  }
}
