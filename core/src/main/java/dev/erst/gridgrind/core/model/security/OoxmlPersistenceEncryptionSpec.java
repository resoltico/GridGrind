package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Explicit encryption disposition for one persisted OOXML package. */
@Union(discriminator = "type", catalog = true, displayOrder = 45)
public sealed interface OoxmlPersistenceEncryptionSpec
    permits dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.None,
        dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.Encrypt,
        dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.PreserveSource {
  /** Deliberately persist plaintext. */
  @Variant(id = "NONE", summary = "Deliberately persist plaintext.")
  record None() implements OoxmlPersistenceEncryptionSpec, ResolvedPersistenceEncryption {}

  /** Encrypt with one explicitly supplied strong OOXML write envelope. */
  @Variant(
      id = "ENCRYPT",
      summary = "Persist with one explicitly supplied OOXML encryption envelope.")
  record Encrypt(OoxmlEncryptionSpec encryption) implements OoxmlPersistenceEncryptionSpec {
    public Encrypt {
      Objects.requireNonNull(encryption, "encryption must not be null");
    }
  }

  /** Mirror the opened source workbook's encryption state. */
  @Variant(
      id = "PRESERVE_SOURCE",
      summary =
          "Reapply the verified source encryption envelope when it is compatible with the AGILE write contract.")
  record PreserveSource()
      implements OoxmlPersistenceEncryptionSpec, ResolvedPersistenceEncryption {}
}
