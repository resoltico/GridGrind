package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.Objects;

/** Factual table mapping entry linked to one custom-XML map. */
public record CustomXmlLinkedTableReport(
    SheetLabel sheetName,
    String tableName,
    String tableDisplayName,
    String range,
    String commonXPath) {
  public CustomXmlLinkedTableReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
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
