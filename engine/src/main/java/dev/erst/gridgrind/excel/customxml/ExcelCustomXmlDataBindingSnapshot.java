package dev.erst.gridgrind.excel.customxml;

import java.util.Objects;
import java.util.Optional;

/** Immutable factual workbook custom-XML data-binding metadata. */
public record ExcelCustomXmlDataBindingSnapshot(
    Optional<String> dataBindingName,
    Optional<Boolean> fileBinding,
    Optional<Long> connectionId,
    Optional<String> fileBindingName,
    long loadMode) {
  public ExcelCustomXmlDataBindingSnapshot {
    dataBindingName = requireNonBlankOptional(dataBindingName, "dataBindingName");
    Objects.requireNonNull(fileBinding, "fileBinding must not be null");
    Objects.requireNonNull(connectionId, "connectionId must not be null");
    if (connectionId.isPresent() && connectionId.orElseThrow() < 0L) {
      throw new IllegalArgumentException("connectionId must not be negative");
    }
    fileBindingName = requireNonBlankOptional(fileBindingName, "fileBindingName");
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

  private static Optional<String> requireNonBlankOptional(
      Optional<String> value, String fieldName) {
    Optional<String> required = Objects.requireNonNull(value, fieldName + " must not be null");
    required.ifPresent(nonBlank -> requireNonBlank(nonBlank, fieldName));
    return required;
  }
}
