package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.ExcelOoxmlEncryptionMode;
import java.util.Objects;

/** OOXML package-encryption settings applied during workbook persistence. */
public record OoxmlEncryptionInput(String password, ExcelOoxmlEncryptionMode mode) {
  public OoxmlEncryptionInput {
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
