package dev.erst.gridgrind.excel.ooxml;

import java.util.Objects;

/** Explicit encryption disposition for one persisted OOXML package. */
public sealed interface ExcelOoxmlPersistenceEncryption
    permits ExcelOoxmlPersistenceEncryption.Plaintext,
        ExcelOoxmlPersistenceEncryption.Encrypt,
        ExcelOoxmlPersistenceEncryption.PreserveSource {
  /** Deliberately persist an unencrypted package. */
  record Plaintext() implements ExcelOoxmlPersistenceEncryption {}

  /** Persist using one explicit OOXML encryption envelope. */
  record Encrypt(ExcelOoxmlEncryptionOptions options) implements ExcelOoxmlPersistenceEncryption {
    public Encrypt {
      Objects.requireNonNull(options, "options must not be null");
    }
  }

  /** Reapply a write-compatible encryption envelope from the opened source package. */
  record PreserveSource() implements ExcelOoxmlPersistenceEncryption {}
}
