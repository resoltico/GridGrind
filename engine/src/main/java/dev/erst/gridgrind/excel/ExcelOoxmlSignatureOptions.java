package dev.erst.gridgrind.excel;

import java.nio.file.Path;
import java.util.Objects;

/** OOXML package-signing settings applied during workbook persistence. */
public record ExcelOoxmlSignatureOptions(
    Path pkcs12Path,
    String keystorePassword,
    String keyPassword,
    String alias,
    ExcelOoxmlSignatureDigestAlgorithm digestAlgorithm,
    String description) {
  public ExcelOoxmlSignatureOptions {
    Objects.requireNonNull(pkcs12Path, "pkcs12Path must not be null");
    keystorePassword = normalizeRequired(keystorePassword, "keystorePassword");
    keyPassword =
        keyPassword == null ? keystorePassword : normalizeRequired(keyPassword, "keyPassword");
    alias = normalizeOptional(alias, "alias");
    digestAlgorithm =
        Objects.requireNonNullElse(digestAlgorithm, ExcelOoxmlSignatureDigestAlgorithm.SHA256);
    description = normalizeOptional(description, "description");
  }

  private static String normalizeRequired(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static String normalizeOptional(String value, String fieldName) {
    if (value == null) {
      return null;
    }
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
