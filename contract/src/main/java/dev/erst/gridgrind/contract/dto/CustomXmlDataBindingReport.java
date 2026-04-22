package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Factual workbook custom-XML data-binding metadata returned by inspection. */
public record CustomXmlDataBindingReport(
    String dataBindingName,
    Boolean fileBinding,
    Long connectionId,
    String fileBindingName,
    long loadMode) {
  public CustomXmlDataBindingReport {
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
