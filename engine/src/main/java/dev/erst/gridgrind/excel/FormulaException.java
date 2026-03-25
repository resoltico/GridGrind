package dev.erst.gridgrind.excel;

/**
 * Sealed marker interface for formula problems detected while writing or evaluating workbook
 * formulas. Each permitted subtype is a concrete exception class carrying the sheet, address, and
 * formula context fields directly.
 */
public sealed interface FormulaException
    permits InvalidFormulaException, UnsupportedFormulaException {
  /** Name of the sheet containing the failing formula cell. */
  String sheetName();

  /** A1-notation address of the failing formula cell. */
  String address();

  /** Raw formula text that triggered the failure. */
  String formula();
}
