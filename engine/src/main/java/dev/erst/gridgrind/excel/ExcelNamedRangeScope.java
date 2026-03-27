package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable workbook-core scope of a defined name, either workbook-wide or sheet-local. */
public sealed interface ExcelNamedRangeScope
    permits ExcelNamedRangeScope.WorkbookScope, ExcelNamedRangeScope.SheetScope {

  /** Workbook-wide named-range scope. */
  record WorkbookScope() implements ExcelNamedRangeScope {}

  /** Sheet-local named-range scope bound to one sheet name. */
  record SheetScope(String sheetName) implements ExcelNamedRangeScope {
    public SheetScope {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }
}
