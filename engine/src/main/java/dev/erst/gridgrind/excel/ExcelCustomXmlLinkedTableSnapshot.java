package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable factual table mapping entry linked to one custom-XML map. */
public record ExcelCustomXmlLinkedTableSnapshot(
    String sheetName, String tableName, String tableDisplayName, String range, String commonXPath) {
  public ExcelCustomXmlLinkedTableSnapshot {
    sheetName = requireNonBlank(sheetName, "sheetName");
    tableName = requireNonBlank(tableName, "tableName");
    tableDisplayName = requireNonBlank(tableDisplayName, "tableDisplayName");
    range = requireNonBlank(range, "range");
    commonXPath = requireNonBlank(commonXPath, "commonXPath");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
