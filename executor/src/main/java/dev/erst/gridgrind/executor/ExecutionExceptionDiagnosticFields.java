package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.FormulaException;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;

/** Exception diagnostic field extraction for execution journal contexts. */
final class ExecutionExceptionDiagnosticFields {
  private ExecutionExceptionDiagnosticFields() {}

  static String sheetNameFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> formulaException.sheetName();
      case SheetNotFoundException sheetNotFoundException -> sheetNotFoundException.sheetName();
      case NamedRangeNotFoundException namedRangeNotFoundException ->
          switch (namedRangeNotFoundException.scope()) {
            case ExcelNamedRangeScope.WorkbookScope _ -> null;
            case ExcelNamedRangeScope.SheetScope sheetScope -> sheetScope.sheetName();
          };
      default -> null;
    };
  }

  static String addressFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> formulaException.address();
      case CellNotFoundException cellNotFoundException -> cellNotFoundException.address();
      case InvalidCellAddressException invalidCellAddressException ->
          invalidCellAddressException.address();
      default -> null;
    };
  }

  static String rangeFor(Exception exception) {
    if (exception instanceof InvalidRangeAddressException invalidRangeAddressException) {
      return invalidRangeAddressException.range();
    }
    return null;
  }

  static String formulaFor(Exception exception) {
    if (exception instanceof FormulaException formulaException) {
      return formulaException.formula();
    }
    return null;
  }

  static String namedRangeNameFor(Exception exception) {
    if (exception instanceof NamedRangeNotFoundException namedRangeNotFoundException) {
      return namedRangeNotFoundException.name();
    }
    return null;
  }
}
