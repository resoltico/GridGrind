package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Total encryption and signature policy for one persisted OOXML package. */
@CatalogType(
    id = "OoxmlPersistenceSecuritySpec",
    summary =
        "Total OOXML encryption and signature policy applied during persistence. It is required for writing EXISTING sources; NEW writes default to encryption NONE and signature NONE when security is omitted.",
    displayOrder = 20)
public record OoxmlPersistenceSecuritySpec(
    OoxmlPersistenceEncryptionSpec encryption, OoxmlPersistenceSignatureSpec signature) {
  public OoxmlPersistenceSecuritySpec {
    java.util.Objects.requireNonNull(encryption, "encryption must not be null");
    java.util.Objects.requireNonNull(signature, "signature must not be null");
  }

  /** Deliberately persists plaintext without a package signature. */
  public static OoxmlPersistenceSecuritySpec none() {
    return new OoxmlPersistenceSecuritySpec(
        new OoxmlPersistenceEncryptionSpec.None(), new OoxmlPersistenceSignatureSpec.None());
  }
}
