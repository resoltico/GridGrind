package dev.erst.gridgrind.excel.ooxml;

import java.util.Objects;
import java.util.Optional;

/** Optional OOXML package-security settings applied during workbook persistence. */
public record ExcelOoxmlPersistenceOptions(
    Optional<ExcelOoxmlEncryptionOptions> encryption,
    Optional<ExcelOoxmlSignatureOptions> signature) {
  public ExcelOoxmlPersistenceOptions {
    Objects.requireNonNull(encryption, "encryption must not be null");
    Objects.requireNonNull(signature, "signature must not be null");
  }

  /** Returns one persistence-options value that applies neither encryption nor signing. */
  public static ExcelOoxmlPersistenceOptions none() {
    return new ExcelOoxmlPersistenceOptions(Optional.empty(), Optional.empty());
  }

  /** Returns whether neither encryption nor signing was requested for persistence. */
  public boolean isEmpty() {
    return encryption.isEmpty() && signature.isEmpty();
  }
}
