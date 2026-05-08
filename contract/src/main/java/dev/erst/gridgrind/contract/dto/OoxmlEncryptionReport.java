package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import java.util.Objects;
import java.util.Optional;

/** Factual OOXML package-encryption report for one workbook package. */
public record OoxmlEncryptionReport(
    boolean encrypted,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ExcelOoxmlEncryptionMode> mode,
    @JsonInclude(JsonInclude.Include.NON_ABSENT)
        Optional<ExcelOoxmlCipherAlgorithm> cipherAlgorithm,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ExcelOoxmlHashAlgorithm> hashAlgorithm,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ExcelOoxmlChainingMode> chainingMode,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> keyBits,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> blockSize,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> spinCount) {
  public OoxmlEncryptionReport {
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
            "Unencrypted package reports must not include encryption detail fields");
      }
    } else {
      requirePresent(mode, "mode");
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

  private static <T> Optional<T> requirePresent(Optional<T> value, String fieldName) {
    Optional<T> required = Objects.requireNonNull(value, fieldName + " must not be null");
    if (required.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be absent");
    }
    return required;
  }
}
