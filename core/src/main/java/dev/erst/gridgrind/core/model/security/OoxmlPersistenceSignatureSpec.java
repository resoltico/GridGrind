package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Explicit signature disposition for one persisted OOXML package. */
@Union(discriminator = "type", catalog = true, displayOrder = 46)
public sealed interface OoxmlPersistenceSignatureSpec
    permits dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.None,
        dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.Sign {
  /** Deliberately persist an unsigned package. */
  @Variant(id = "NONE", summary = "Deliberately persist without an OOXML package signature.")
  record None() implements OoxmlPersistenceSignatureSpec {}

  /** Sign the package with explicitly supplied material. */
  @Variant(id = "SIGN", summary = "Sign with explicitly supplied PKCS#12 material.")
  record Sign(OoxmlSignatureSpec signature) implements OoxmlPersistenceSignatureSpec {
    public Sign {
      Objects.requireNonNull(signature, "signature must not be null");
    }
  }
}
