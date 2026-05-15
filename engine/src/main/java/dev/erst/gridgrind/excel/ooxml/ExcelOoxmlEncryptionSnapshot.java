package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import java.util.Objects;
import java.util.Optional;

/** Immutable factual snapshot of one workbook package's OOXML encryption state. */
@SuppressWarnings("PMD.CommentRequired")
public record ExcelOoxmlEncryptionSnapshot(
    boolean encrypted,
    Optional<ExcelOoxmlEncryptionMode> mode,
    Optional<ExcelOoxmlCipherAlgorithm> cipherAlgorithm,
    Optional<ExcelOoxmlHashAlgorithm> hashAlgorithm,
    Optional<ExcelOoxmlChainingMode> chainingMode,
    Optional<Integer> keyBits,
    Optional<Integer> blockSize,
    Optional<Integer> spinCount) {
  public ExcelOoxmlEncryptionSnapshot {
    Objects.requireNonNull(mode, "mode must not be null");
    Objects.requireNonNull(cipherAlgorithm, "cipherAlgorithm must not be null");
    Objects.requireNonNull(hashAlgorithm, "hashAlgorithm must not be null");
    Objects.requireNonNull(chainingMode, "chainingMode must not be null");
    Objects.requireNonNull(keyBits, "keyBits must not be null");
    Objects.requireNonNull(blockSize, "blockSize must not be null");
    Objects.requireNonNull(spinCount, "spinCount must not be null");
    if (!encrypted) {
      if (mode.isPresent()
          || cipherAlgorithm.isPresent()
          || hashAlgorithm.isPresent()
          || chainingMode.isPresent()
          || keyBits.isPresent()
          || blockSize.isPresent()
          || spinCount.isPresent()) {
        throw new IllegalArgumentException(
            "Unencrypted package snapshots must not include encryption detail fields");
      }
    } else {
      if (mode.isEmpty()) {
        throw new IllegalArgumentException("mode must not be absent when encrypted");
      }
      requirePresent(cipherAlgorithm, "cipherAlgorithm");
      requirePresent(hashAlgorithm, "hashAlgorithm");
      requirePresent(chainingMode, "chainingMode");
      if (keyBits.isEmpty() || keyBits.orElseThrow() <= 0) {
        throw new IllegalArgumentException("keyBits must be positive when encrypted");
      }
      if (blockSize.isEmpty() || blockSize.orElseThrow() <= 0) {
        throw new IllegalArgumentException("blockSize must be positive when encrypted");
      }
      if (spinCount.isEmpty() || spinCount.orElseThrow() < 0) {
        throw new IllegalArgumentException("spinCount must be zero or positive when encrypted");
      }
    }
  }

  public static ExcelOoxmlEncryptionSnapshot none() {
    return new ExcelOoxmlEncryptionSnapshot(
        false,
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static <T> Optional<T> requirePresent(Optional<T> value, String fieldName) {
    Optional<T> required = Objects.requireNonNull(value, fieldName + " must not be null");
    if (required.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be absent");
    }
    return required;
  }
}
