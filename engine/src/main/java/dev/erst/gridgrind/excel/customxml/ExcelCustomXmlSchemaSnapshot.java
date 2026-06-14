package dev.erst.gridgrind.excel.customxml;

import java.util.Objects;
import java.util.Optional;

/** Optional schema metadata attached to one custom-XML mapping. */
public record ExcelCustomXmlSchemaSnapshot(
    Optional<String> namespace,
    Optional<String> language,
    Optional<String> reference,
    Optional<String> xml) {
  public ExcelCustomXmlSchemaSnapshot {
    namespace = requireNonBlankOptional(namespace, "namespace");
    language = requireNonBlankOptional(language, "language");
    reference = requireNonBlankOptional(reference, "reference");
    xml = requireNonBlankOptional(xml, "xml");
  }

  private static Optional<String> requireNonBlankOptional(
      Optional<String> value, String fieldName) {
    Optional<String> required = Objects.requireNonNullElseGet(value, Optional::empty);
    required.ifPresent(nonBlank -> requireNonBlank(nonBlank, fieldName));
    return required;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
