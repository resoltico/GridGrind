package dev.erst.gridgrind.core.model.security;

import java.util.Objects;

/** Explicit encryption disposition for one persisted OOXML package. */
public sealed interface ResolvedPersistenceEncryption
    permits OoxmlPersistenceEncryptionSpec.None,
        ResolvedPersistenceEncryption.Encrypt,
        OoxmlPersistenceEncryptionSpec.PreserveSource {

  /** Persist using one explicit OOXML encryption envelope. */
  record Encrypt(ResolvedEncryption options) implements ResolvedPersistenceEncryption {
    public Encrypt {
      Objects.requireNonNull(options, "options must not be null");
    }
  }
}
