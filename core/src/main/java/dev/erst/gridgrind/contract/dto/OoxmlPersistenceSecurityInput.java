package dev.erst.gridgrind.contract.dto;

/** Total encryption and signature policy for one persisted OOXML package. */
public record OoxmlPersistenceSecurityInput(
    OoxmlPersistenceEncryptionInput encryption, OoxmlPersistenceSignatureInput signature) {
  public OoxmlPersistenceSecurityInput {
    java.util.Objects.requireNonNull(encryption, "encryption must not be null");
    java.util.Objects.requireNonNull(signature, "signature must not be null");
  }

  /** Deliberately persists plaintext without a package signature. */
  public static OoxmlPersistenceSecurityInput none() {
    return new OoxmlPersistenceSecurityInput(
        new OoxmlPersistenceEncryptionInput.None(), new OoxmlPersistenceSignatureInput.None());
  }
}
