package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.FormulaException;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import java.util.Optional;

/** Exception diagnostic field extraction for execution journal contexts. */
final class ExecutionExceptionDiagnosticFields {
  private ExecutionExceptionDiagnosticFields() {}

  static Optional<String> sheetNameFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> Optional.ofNullable(formulaException.sheetName());
      case SheetNotFoundException sheetNotFoundException ->
          Optional.ofNullable(sheetNotFoundException.sheetName());
      case NamedRangeNotFoundException namedRangeNotFoundException ->
          switch (namedRangeNotFoundException.scope()) {
            case ExcelNamedRangeScope.WorkbookScope _ -> Optional.empty();
            case ExcelNamedRangeScope.SheetScope sheetScope ->
                Optional.ofNullable(sheetScope.sheetName());
          };
      default -> Optional.empty();
    };
  }

  static Optional<String> addressFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> Optional.ofNullable(formulaException.address());
      case CellNotFoundException cellNotFoundException ->
          Optional.ofNullable(cellNotFoundException.address());
      case InvalidCellAddressException invalidCellAddressException ->
          Optional.ofNullable(invalidCellAddressException.address());
      default -> Optional.empty();
    };
  }

  static Optional<String> rangeFor(Exception exception) {
    if (exception instanceof InvalidRangeAddressException invalidRangeAddressException) {
      return Optional.ofNullable(invalidRangeAddressException.range());
    }
    return Optional.empty();
  }

  static Optional<String> formulaFor(Exception exception) {
    if (exception instanceof FormulaException formulaException) {
      return Optional.ofNullable(formulaException.formula());
    }
    return Optional.empty();
  }

  static Optional<String> namedRangeNameFor(Exception exception) {
    if (exception instanceof NamedRangeNotFoundException namedRangeNotFoundException) {
      return Optional.ofNullable(namedRangeNotFoundException.name());
    }
    return Optional.empty();
  }
}
