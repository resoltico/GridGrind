package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Cell-value commands that write, clear, and append rectangular content. */
public sealed interface WorkbookCellCommand extends WorkbookCommand
    permits WorkbookCellCommand.SetCell,
        WorkbookCellCommand.SetRange,
        WorkbookCellCommand.ClearRange,
        WorkbookCellCommand.SetArrayFormula,
        WorkbookCellCommand.ClearArrayFormula,
        WorkbookCellCommand.AppendRow {

  record SetCell(String sheetName, String address, ExcelCellValue value)
      implements WorkbookCellCommand {
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
      implements WorkbookCellCommand {
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

  record ClearRange(String sheetName, String range) implements WorkbookCellCommand {
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

  record SetArrayFormula(String sheetName, String range, ExcelArrayFormulaDefinition formula)
      implements WorkbookCellCommand {
    public SetArrayFormula {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(formula, "formula must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (range.isBlank()) {
        throw new IllegalArgumentException("range must not be blank");
      }
    }
  }

  record ClearArrayFormula(String sheetName, String address) implements WorkbookCellCommand {
    public ClearArrayFormula {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(address, "address must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (address.isBlank()) {
        throw new IllegalArgumentException("address must not be blank");
      }
    }
  }

  record AppendRow(String sheetName, List<ExcelCellValue> values) implements WorkbookCellCommand {
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
}
