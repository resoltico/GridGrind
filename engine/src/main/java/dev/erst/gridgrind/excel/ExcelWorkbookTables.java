package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Workbook-global table authoring operations. */
public final class ExcelWorkbookTables {
  private final ExcelWorkbook workbook;

  ExcelWorkbookTables(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Creates or replaces one workbook-global table definition. */
  public ExcelWorkbook setTable(ExcelTableDefinition definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    new ExcelTableController().setTable(workbook, definition);
    return workbook;
  }

  /** Deletes one existing table by workbook-global name and expected sheet name. */
  public ExcelWorkbook deleteTable(String name, String sheetName) {
    new ExcelTableController().deleteTable(workbook, name, sheetName);
    return workbook;
  }
}
