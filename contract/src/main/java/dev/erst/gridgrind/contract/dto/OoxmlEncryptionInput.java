package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import java.util.Objects;

/** OOXML package-encryption settings applied during workbook persistence. */
public record OoxmlEncryptionInput(String password, ExcelOoxmlEncryptionMode mode) {
  public OoxmlEncryptionInput {
    password = normalizeRequired(password, "password");
    Objects.requireNonNull(mode, "mode must not be null");
  }

  /** Creates one AGILE OOXML encryption payload explicitly. */
  public static OoxmlEncryptionInput agile(String password) {
    return new OoxmlEncryptionInput(password, ExcelOoxmlEncryptionMode.AGILE);
  }

  private static String normalizeRequired(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
