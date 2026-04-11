package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelSheetNames;
import java.util.Objects;

/** Protocol-facing table definition attached to one table authoring request. */
public record TableInput(
    String name, String sheetName, String range, Boolean showTotalsRow, TableStyleInput style) {
  public TableInput {
    name = ProtocolDefinedNameValidation.validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    showTotalsRow = Boolean.TRUE.equals(showTotalsRow);
    Objects.requireNonNull(style, "style must not be null");
  }
}
