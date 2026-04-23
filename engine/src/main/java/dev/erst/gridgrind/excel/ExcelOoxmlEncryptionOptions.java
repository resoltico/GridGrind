package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import java.util.Objects;

/** OOXML package-encryption settings applied during workbook persistence. */
public record ExcelOoxmlEncryptionOptions(String password, ExcelOoxmlEncryptionMode mode) {
  public ExcelOoxmlEncryptionOptions {
    password = normalizeRequired(password, "password");
    mode = Objects.requireNonNullElse(mode, ExcelOoxmlEncryptionMode.AGILE);
  }

  private static String normalizeRequired(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
