package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import java.util.Objects;

/** Immutable factual snapshot of one workbook package's OOXML encryption state. */
public record ExcelOoxmlEncryptionSnapshot(
    boolean encrypted,
    ExcelOoxmlEncryptionMode mode,
    String cipherAlgorithm,
    String hashAlgorithm,
    String chainingMode,
    Integer keyBits,
    Integer blockSize,
    Integer spinCount) {
  public ExcelOoxmlEncryptionSnapshot {
    if (!encrypted) {
      if (mode != null
          || cipherAlgorithm != null
          || hashAlgorithm != null
          || chainingMode != null
          || keyBits != null
          || blockSize != null
          || spinCount != null) {
        throw new IllegalArgumentException(
            "Unencrypted package snapshots must not include encryption detail fields");
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

  static ExcelOoxmlEncryptionSnapshot none() {
    return new ExcelOoxmlEncryptionSnapshot(false, null, null, null, null, null, null, null);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
