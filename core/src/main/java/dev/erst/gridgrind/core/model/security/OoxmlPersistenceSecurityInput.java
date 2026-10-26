package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Total encryption and signature policy for one persisted OOXML package. */
@CatalogType(
    id = "OoxmlPersistenceSecurityInput",
    summary =
        "Total OOXML encryption and signature policy applied during persistence. It is required for writing EXISTING sources; NEW writes default to encryption NONE and signature NONE when security is omitted.",
    displayOrder = 20)
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
