package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Table-, pivot-, and autofilter-oriented tabular commands. */
public sealed interface WorkbookTabularCommand extends WorkbookCommand
    permits WorkbookTabularCommand.SetPivotTable,
        WorkbookTabularCommand.SetAutofilter,
        WorkbookTabularCommand.ClearAutofilter,
        WorkbookTabularCommand.SetTable,
        WorkbookTabularCommand.DeleteTable,
        WorkbookTabularCommand.DeletePivotTable {

  /** Creates or replaces one workbook-global pivot-table definition. */
  record SetPivotTable(ExcelPivotTableDefinition definition) implements WorkbookTabularCommand {
    public SetPivotTable {
      Objects.requireNonNull(definition, "definition must not be null");
    }
  }

  /** Creates or replaces one sheet-level autofilter range. */
  record SetAutofilter(
      String sheetName,
      String range,
      List<ExcelAutofilterFilterColumn> criteria,
      ExcelAutofilterSortState sortState)
      implements WorkbookTabularCommand {
    /** Creates a plain sheet-level autofilter without criteria or explicit sort state. */
    public SetAutofilter(String sheetName, String range) {
      this(sheetName, range, List.of(), null);
    }

    public SetAutofilter {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      criteria = List.copyOf(Objects.requireNonNull(criteria, "criteria must not be null"));
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
      for (ExcelAutofilterFilterColumn criterion : criteria) {
        Objects.requireNonNull(criterion, "criteria must not contain null values");
      }
    }
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  record ClearAutofilter(String sheetName) implements WorkbookTabularCommand {
    public ClearAutofilter {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one workbook-global table definition. */
  record SetTable(ExcelTableDefinition definition) implements WorkbookTabularCommand {
    public SetTable {
      Objects.requireNonNull(definition, "definition must not be null");
    }
  }

  /** Deletes one existing table by workbook-global name and expected sheet. */
  record DeleteTable(String name, String sheetName) implements WorkbookTabularCommand {
    public DeleteTable {
      name = ExcelTableDefinition.validateName(name);
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Deletes one existing pivot table by workbook-global name and expected sheet. */
  record DeletePivotTable(String name, String sheetName) implements WorkbookTabularCommand {
    public DeletePivotTable {
      name = ExcelPivotTableDefinition.validateName(name);
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }
}
