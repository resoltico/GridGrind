package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Factual table mapping entry linked to one custom-XML map. */
public record CustomXmlLinkedTableReport(
    String sheetName, String tableName, String tableDisplayName, String range, String commonXPath) {
  public CustomXmlLinkedTableReport {
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
