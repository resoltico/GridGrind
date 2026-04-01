package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelTableDefinition;
import java.util.Objects;

/** Protocol-facing table definition attached to one table authoring request. */
public record TableInput(
    String name, String sheetName, String range, boolean showTotalsRow, TableStyleInput style) {
  public TableInput {
    name = dev.erst.gridgrind.excel.ExcelTableDefinition.validateName(name);
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    if (sheetName.length() > 31) {
      throw new IllegalArgumentException("sheetName must not exceed 31 characters: " + sheetName);
    }
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Objects.requireNonNull(style, "style must not be null");
  }

  /** Converts this protocol payload into the workbook-core table definition. */
  public ExcelTableDefinition toExcelTableDefinition() {
    return new ExcelTableDefinition(
        name, sheetName, range, showTotalsRow, style.toExcelTableStyle());
  }
}
