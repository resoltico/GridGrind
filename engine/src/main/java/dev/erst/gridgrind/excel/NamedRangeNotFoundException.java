package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Signals that a requested defined name does not exist in the requested workbook scope. */
public final class NamedRangeNotFoundException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final String name;
  private final ExcelNamedRangeScope scope;

  /** Creates the exception for the missing defined name and scope. */
  public NamedRangeNotFoundException(String name, ExcelNamedRangeScope scope) {
    super("Named range not found: " + describe(name, scope));
    this.name = Objects.requireNonNull(name, "name must not be null");
    this.scope = Objects.requireNonNull(scope, "scope must not be null");
  }

  /** Returns the missing defined-name identifier. */
  public String name() {
    return name;
  }

  /** Returns the workbook or sheet scope in which the lookup failed. */
  public ExcelNamedRangeScope scope() {
    return scope;
  }

  private static String describe(String name, ExcelNamedRangeScope scope) {
    return switch (Objects.requireNonNull(scope, "scope must not be null")) {
      case ExcelNamedRangeScope.WorkbookScope _ -> name + " (workbook scope)";
      case ExcelNamedRangeScope.SheetScope sheetScope ->
          name + " (sheet scope: " + sheetScope.sheetName() + ")";
    };
  }
}
