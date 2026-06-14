package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import java.util.Objects;

/** Immutable factual snapshot of one workbook package's OOXML encryption state. */
@SuppressWarnings("PMD.CommentRequired")
public sealed interface ExcelOoxmlEncryptionSnapshot
    permits ExcelOoxmlEncryptionSnapshot.None, ExcelOoxmlEncryptionSnapshot.Encrypted {
  static None none() {
    return new None();
  }

  record None() implements ExcelOoxmlEncryptionSnapshot {}

  record Encrypted(
      ExcelOoxmlEncryptionMode mode,
      ExcelOoxmlCipherAlgorithm cipherAlgorithm,
      ExcelOoxmlHashAlgorithm hashAlgorithm,
      ExcelOoxmlChainingMode chainingMode,
      int keyBits,
      int blockSize,
      int spinCount)
      implements ExcelOoxmlEncryptionSnapshot {
    public Encrypted {
      Objects.requireNonNull(mode, "mode must not be null");
      Objects.requireNonNull(cipherAlgorithm, "cipherAlgorithm must not be null");
      Objects.requireNonNull(hashAlgorithm, "hashAlgorithm must not be null");
      Objects.requireNonNull(chainingMode, "chainingMode must not be null");
      if (keyBits <= 0) {
        throw new IllegalArgumentException("keyBits must be positive");
      }
      if (blockSize <= 0) {
        throw new IllegalArgumentException("blockSize must be positive");
      }
      if (spinCount < 0) {
        throw new IllegalArgumentException("spinCount must be zero or positive");
      }
    }
  }
}
