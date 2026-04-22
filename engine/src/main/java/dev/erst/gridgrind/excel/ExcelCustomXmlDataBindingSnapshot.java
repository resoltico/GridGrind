package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable factual workbook custom-XML data-binding metadata. */
public record ExcelCustomXmlDataBindingSnapshot(
    String dataBindingName,
    Boolean fileBinding,
    Long connectionId,
    String fileBindingName,
    long loadMode) {
  public ExcelCustomXmlDataBindingSnapshot {
    if (dataBindingName != null) {
      dataBindingName = requireNonBlank(dataBindingName, "dataBindingName");
    }
    if (connectionId != null && connectionId < 0L) {
      throw new IllegalArgumentException("connectionId must not be negative");
    }
    if (fileBindingName != null) {
      fileBindingName = requireNonBlank(fileBindingName, "fileBindingName");
    }
    if (loadMode < 0L) {
      throw new IllegalArgumentException("loadMode must not be negative");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
