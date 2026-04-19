package dev.erst.gridgrind.contract.dto;

/** Optional OOXML package-encryption and package-signing settings for workbook persistence. */
public record OoxmlPersistenceSecurityInput(
    OoxmlEncryptionInput encryption, OoxmlSignatureInput signature) {
  public OoxmlPersistenceSecurityInput {
    if (encryption == null && signature == null) {
      throw new IllegalArgumentException(
          "At least one of encryption or signature must be supplied when security is present");
    }
  }
}
