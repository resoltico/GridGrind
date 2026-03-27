package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Exact selector used to match named ranges inside workbook-core read commands. */
public sealed interface ExcelNamedRangeSelector
    permits ExcelNamedRangeSelector.ByName,
        ExcelNamedRangeSelector.WorkbookScope,
        ExcelNamedRangeSelector.SheetScope {

  /** Matches named ranges by identifier across all scopes. */
  record ByName(String name) implements ExcelNamedRangeSelector {
    public ByName {
      name = requireNonBlank(name, "name");
    }
  }

  /** Matches the workbook-scoped named range with the given identifier. */
  record WorkbookScope(String name) implements ExcelNamedRangeSelector {
    public WorkbookScope {
      name = requireNonBlank(name, "name");
    }
  }

  /** Matches the sheet-scoped named range with the given identifier on one sheet. */
  record SheetScope(String name, String sheetName) implements ExcelNamedRangeSelector {
    public SheetScope {
      name = requireNonBlank(name, "name");
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
