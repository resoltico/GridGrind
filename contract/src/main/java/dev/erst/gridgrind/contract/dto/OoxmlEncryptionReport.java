package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import java.util.Objects;

/** Factual OOXML package-encryption report for one workbook package. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = OoxmlEncryptionReport.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = OoxmlEncryptionReport.Encrypted.class, name = "ENCRYPTED")
})
public sealed interface OoxmlEncryptionReport
    permits OoxmlEncryptionReport.None, OoxmlEncryptionReport.Encrypted {
  /** Package carries no OOXML encryption envelope. */
  record None() implements OoxmlEncryptionReport {}

  /** Package is encrypted with one fully specified OOXML encryption envelope. */
  record Encrypted(
      ExcelOoxmlEncryptionMode mode,
      ExcelOoxmlCipherAlgorithm cipherAlgorithm,
      ExcelOoxmlHashAlgorithm hashAlgorithm,
      ExcelOoxmlChainingMode chainingMode,
      int keyBits,
      int blockSize,
      int spinCount)
      implements OoxmlEncryptionReport {
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
