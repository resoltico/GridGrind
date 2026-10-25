package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Explicit encryption disposition for one persisted OOXML package. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = OoxmlPersistenceEncryptionInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = OoxmlPersistenceEncryptionInput.Encrypt.class, name = "ENCRYPT"),
  @JsonSubTypes.Type(
      value = OoxmlPersistenceEncryptionInput.PreserveSource.class,
      name = "PRESERVE_SOURCE")
})
public sealed interface OoxmlPersistenceEncryptionInput
    permits OoxmlPersistenceEncryptionInput.None,
        OoxmlPersistenceEncryptionInput.Encrypt,
        OoxmlPersistenceEncryptionInput.PreserveSource {
  /** Deliberately persist plaintext. */
  record None() implements OoxmlPersistenceEncryptionInput {}

  /** Encrypt with one explicitly supplied strong OOXML write envelope. */
  record Encrypt(OoxmlEncryptionInput encryption) implements OoxmlPersistenceEncryptionInput {
    public Encrypt {
      Objects.requireNonNull(encryption, "encryption must not be null");
    }
  }

  /** Mirror the opened source workbook's encryption state. */
  record PreserveSource() implements OoxmlPersistenceEncryptionInput {}
}
