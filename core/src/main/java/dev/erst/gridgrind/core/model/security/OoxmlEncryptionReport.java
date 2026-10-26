package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import java.util.Objects;

/** Factual OOXML package-encryption report for one workbook package. */
@Union(discriminator = "type", catalog = true, displayOrder = 44)
public sealed interface OoxmlEncryptionReport
    permits dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport.None,
        dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport.Encrypted {
  /** Package carries no OOXML encryption envelope. */
  @Variant(id = "NONE", summary = "Package carries no OOXML encryption envelope.")
  record None() implements OoxmlEncryptionReport {}

  /** Package is encrypted with one fully specified OOXML encryption envelope. */
  @Variant(
      id = "ENCRYPTED",
      summary = "Package is encrypted with one fully specified OOXML encryption envelope.")
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
