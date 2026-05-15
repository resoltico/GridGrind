package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** OOXML package-signing settings applied during workbook persistence. */
public record ExcelOoxmlSignatureOptions(
    Path pkcs12Path,
    String keystorePassword,
    String keyPassword,
    @Nullable String alias,
    ExcelOoxmlSignatureDigestAlgorithm digestAlgorithm,
    @Nullable String description) {
  public ExcelOoxmlSignatureOptions {
    Objects.requireNonNull(pkcs12Path, "pkcs12Path must not be null");
    keystorePassword = normalizeRequired(keystorePassword, "keystorePassword");
    keyPassword =
        keyPassword == null ? keystorePassword : normalizeRequired(keyPassword, "keyPassword");
    alias = normalizeOptional(alias, "alias").orElse(null);
    digestAlgorithm =
        Objects.requireNonNullElse(digestAlgorithm, ExcelOoxmlSignatureDigestAlgorithm.SHA256);
    description = normalizeOptional(description, "description").orElse(null);
  }

  private static String normalizeRequired(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static Optional<String> normalizeOptional(@Nullable String value, String fieldName) {
    if (value == null) {
      return Optional.empty();
    }
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return Optional.of(value);
  }
}
