package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableController;
import java.util.Objects;

/** Workbook-global pivot-table authoring operations. */
public final class ExcelWorkbookPivots {
  private final ExcelWorkbook workbook;

  ExcelWorkbookPivots(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Creates or replaces one workbook-global pivot-table definition. */
  public ExcelWorkbook setPivotTable(PivotTableSpec definition) {
    Objects.requireNonNull(definition, "definition must not be null");
    new ExcelPivotTableController().setPivotTable(workbook, definition);
    return workbook;
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet name. */
  public ExcelWorkbook deletePivotTable(String name, String sheetName) {
    new ExcelPivotTableController().deletePivotTable(workbook, name, sheetName);
    return workbook;
  }
}
