package dev.erst.gridgrind.core.model.security;

import java.util.Objects;

/** Explicit package-signature disposition for one persisted OOXML package. */
public sealed interface ResolvedPersistenceSignature
    permits OoxmlPersistenceSignatureSpec.None, ResolvedPersistenceSignature.Sign {

  /** Persist with one explicit OOXML package signature. */
  record Sign(ResolvedSignature options) implements ResolvedPersistenceSignature {
    public Sign {
      Objects.requireNonNull(options, "options must not be null");
    }
  }
}
