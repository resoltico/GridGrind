package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Explicit signature disposition for one persisted OOXML package. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = OoxmlPersistenceSignatureInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = OoxmlPersistenceSignatureInput.Sign.class, name = "SIGN")
})
public sealed interface OoxmlPersistenceSignatureInput
    permits OoxmlPersistenceSignatureInput.None, OoxmlPersistenceSignatureInput.Sign {
  /** Deliberately persist an unsigned package. */
  record None() implements OoxmlPersistenceSignatureInput {}

  /** Sign the package with explicitly supplied material. */
  record Sign(OoxmlSignatureInput signature) implements OoxmlPersistenceSignatureInput {
    public Sign {
      Objects.requireNonNull(signature, "signature must not be null");
    }
  }
}
