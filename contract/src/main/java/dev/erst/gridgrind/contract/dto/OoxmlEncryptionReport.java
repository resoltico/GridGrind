package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.ExcelOoxmlEncryptionMode;
import java.util.Objects;

/** Factual OOXML package-encryption report for one workbook package. */
public record OoxmlEncryptionReport(
    boolean encrypted,
    ExcelOoxmlEncryptionMode mode,
    String cipherAlgorithm,
    String hashAlgorithm,
    String chainingMode,
    Integer keyBits,
    Integer blockSize,
    Integer spinCount) {
  public OoxmlEncryptionReport {
    if (!encrypted) {
      if (mode != null
          || cipherAlgorithm != null
          || hashAlgorithm != null
          || chainingMode != null
          || keyBits != null
          || blockSize != null
          || spinCount != null) {
        throw new IllegalArgumentException(
            "Unencrypted package reports must not include encryption detail fields");
      }
    } else {
      Objects.requireNonNull(mode, "mode must not be null when encrypted");
      cipherAlgorithm = requireNonBlank(cipherAlgorithm, "cipherAlgorithm");
      hashAlgorithm = requireNonBlank(hashAlgorithm, "hashAlgorithm");
      chainingMode = requireNonBlank(chainingMode, "chainingMode");
      if (keyBits == null || keyBits <= 0) {
        throw new IllegalArgumentException("keyBits must be positive when encrypted");
      }
      if (blockSize == null || blockSize <= 0) {
        throw new IllegalArgumentException("blockSize must be positive when encrypted");
      }
      if (spinCount == null || spinCount < 0) {
        throw new IllegalArgumentException("spinCount must be zero or positive when encrypted");
      }
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
