package dev.erst.gridgrind.core.model.security;

import java.util.Objects;

/** Total OOXML package-security policy applied during workbook persistence. */
public record ResolvedPersistenceSecurity(
    ResolvedPersistenceEncryption encryption, ResolvedPersistenceSignature signature) {
  public ResolvedPersistenceSecurity {
    Objects.requireNonNull(encryption, "encryption must not be null");
    Objects.requireNonNull(signature, "signature must not be null");
  }

  /** Returns the explicit plaintext and unsigned output policy. */
  public static ResolvedPersistenceSecurity none() {
    return new ResolvedPersistenceSecurity(
        new OoxmlPersistenceEncryptionSpec.None(), new OoxmlPersistenceSignatureSpec.None());
  }

  /** Returns whether the policy deliberately writes plaintext without a package signature. */
  public boolean writesPlaintextUnsigned() {
    return encryption instanceof OoxmlPersistenceEncryptionSpec.None
        && signature instanceof OoxmlPersistenceSignatureSpec.None;
  }
}
