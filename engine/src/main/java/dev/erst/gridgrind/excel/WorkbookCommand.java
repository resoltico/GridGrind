package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/**
 * Workbook-core commands that can be executed deterministically against an {@link ExcelWorkbook}.
 */
public sealed interface WorkbookCommand
    permits WorkbookCommand.CreateSheet,
        WorkbookCommand.SetCell,
        WorkbookCommand.SetRange,
        WorkbookCommand.ClearRange,
        WorkbookCommand.ApplyStyle,
        WorkbookCommand.AppendRow,
        WorkbookCommand.AutoSizeColumns,
        WorkbookCommand.EvaluateAllFormulas,
        WorkbookCommand.ForceFormulaRecalculationOnOpen {

  record CreateSheet(String sheetName) implements WorkbookCommand {
    public CreateSheet {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  record SetCell(String sheetName, String address, ExcelCellValue value)
      implements WorkbookCommand {
    public SetCell {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      Objects.requireNonNull(value, "value must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  record SetRange(String sheetName, String range, List<List<ExcelCellValue>> rows)
      implements WorkbookCommand {
    public SetRange {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(rows, "rows must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
      rows = List.copyOf(rows);
      if (rows.isEmpty()) {
        throw new IllegalArgumentException("rows must not be empty");
      }
      int expectedWidth = -1;
      for (List<ExcelCellValue> row : rows) {
        Objects.requireNonNull(row, "rows must not contain nulls");
        List<ExcelCellValue> copiedRow = List.copyOf(row);
        if (copiedRow.isEmpty()) {
          throw new IllegalArgumentException("rows must not contain empty rows");
        }
        if (expectedWidth < 0) {
          expectedWidth = copiedRow.size();
        } else if (copiedRow.size() != expectedWidth) {
          throw new IllegalArgumentException("rows must describe a rectangular matrix");
        }
        for (ExcelCellValue value : copiedRow) {
          Objects.requireNonNull(value, "rows must not contain null cell values");
        }
      }
    }
  }

  record ClearRange(String sheetName, String range) implements WorkbookCommand {
    public ClearRange {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  record ApplyStyle(String sheetName, String range, ExcelCellStyle style)
      implements WorkbookCommand {
    public ApplyStyle {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(style, "style must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  record AppendRow(String sheetName, List<ExcelCellValue> values) implements WorkbookCommand {
    public AppendRow {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(values, "values must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      values = List.copyOf(values);
      if (values.isEmpty()) {
        throw new IllegalArgumentException("values must not be empty");
      }
      for (ExcelCellValue value : values) {
        Objects.requireNonNull(value, "value must not be null");
      }
    }
  }

  record AutoSizeColumns(String sheetName, List<String> columnNames) implements WorkbookCommand {
    public AutoSizeColumns {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(columnNames, "columnNames must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      columnNames = List.copyOf(columnNames);
      if (columnNames.isEmpty()) {
        throw new IllegalArgumentException("columnNames must not be empty");
      }
      for (String columnName : columnNames) {
        Objects.requireNonNull(columnName, "columnName must not be null");
        if (columnName.isBlank()) {
          throw new IllegalArgumentException("columnName must not be blank");
        }
      }
    }
  }

  record EvaluateAllFormulas() implements WorkbookCommand {}

  record ForceFormulaRecalculationOnOpen() implements WorkbookCommand {}
}
