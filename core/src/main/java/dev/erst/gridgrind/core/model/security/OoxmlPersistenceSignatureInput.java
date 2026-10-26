package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Explicit signature disposition for one persisted OOXML package. */
@Union(discriminator = "type", catalog = true, displayOrder = 46)
public sealed interface OoxmlPersistenceSignatureInput
    permits dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureInput.None,
        dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureInput.Sign {
  /** Deliberately persist an unsigned package. */
  @Variant(id = "NONE", summary = "Deliberately persist without an OOXML package signature.")
  record None() implements OoxmlPersistenceSignatureInput {}

  /** Sign the package with explicitly supplied material. */
  @Variant(id = "SIGN", summary = "Sign with explicitly supplied PKCS#12 material.")
  record Sign(OoxmlSignatureInput signature) implements OoxmlPersistenceSignatureInput {
    public Sign {
      Objects.requireNonNull(signature, "signature must not be null");
    }
  }
}
