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
      case FormulaException formulaException -> Optional.of(formulaException.sheetName());
      case SheetNotFoundException sheetNotFoundException ->
          Optional.of(sheetNotFoundException.sheetName());
      case NamedRangeNotFoundException namedRangeNotFoundException ->
          switch (namedRangeNotFoundException.scope()) {
            case ExcelNamedRangeScope.WorkbookScope _ -> Optional.empty();
            case ExcelNamedRangeScope.SheetScope sheetScope -> Optional.of(sheetScope.sheetName());
          };
      default -> Optional.empty();
    };
  }

  static Optional<String> addressFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> Optional.of(formulaException.address());
      case CellNotFoundException cellNotFoundException ->
          Optional.of(cellNotFoundException.address());
      case InvalidCellAddressException invalidCellAddressException ->
          Optional.of(invalidCellAddressException.address());
      default -> Optional.empty();
    };
  }

  static Optional<String> rangeFor(Exception exception) {
    if (exception instanceof InvalidRangeAddressException invalidRangeAddressException) {
      return Optional.of(invalidRangeAddressException.range());
    }
    return Optional.empty();
  }

  static Optional<String> formulaFor(Exception exception) {
    if (exception instanceof FormulaException formulaException) {
      return Optional.of(formulaException.formula());
    }
    return Optional.empty();
  }

  static Optional<String> namedRangeNameFor(Exception exception) {
    if (exception instanceof NamedRangeNotFoundException namedRangeNotFoundException) {
      return Optional.of(namedRangeNotFoundException.name());
    }
    return Optional.empty();
  }
}
