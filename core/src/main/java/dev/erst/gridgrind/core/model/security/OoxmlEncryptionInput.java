package dev.erst.gridgrind.core.model.security;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import java.util.Objects;

/** OOXML package-encryption settings applied during workbook persistence. */
@CatalogType(
    id = "OoxmlEncryptionInput",
    summary =
        "OOXML package-encryption settings for workbook persistence. GridGrind writes AGILE packages only. cipher defaults to AES_256 and hash defaults to SHA_512 when omitted. Supported ciphers: AES_256 and AES_192. Supported hashes: SHA_512, SHA_384, and SHA_256 (LIM-038).",
    displayOrder = 21)
public record OoxmlEncryptionInput(
    SecretReference passwordRef,
    @ProtocolField(optional = true) ExcelOoxmlWriteCipher cipher,
    @ProtocolField(optional = true) ExcelOoxmlWriteHash hash) {
  /** Reads one OOXML encryption block while applying the documented omission default. */
  @JsonCreator
  static OoxmlEncryptionInput create(
      @JsonProperty("passwordRef") SecretReference passwordRef,
      @JsonProperty("cipher") ExcelOoxmlWriteCipher cipher,
      @JsonProperty("hash") ExcelOoxmlWriteHash hash) { // LIM-038
    return new OoxmlEncryptionInput(passwordRef, normalizeCipher(cipher), normalizeHash(hash));
  }

  public OoxmlEncryptionInput {
    Objects.requireNonNull(passwordRef, "passwordRef must not be null");
    Objects.requireNonNull(cipher, "cipher must not be null");
    Objects.requireNonNull(hash, "hash must not be null");
  }

  /** Creates one strong OOXML encryption payload with GridGrind's default write envelope. */
  public static OoxmlEncryptionInput strong(SecretReference passwordRef) {
    return new OoxmlEncryptionInput(
        passwordRef, ExcelOoxmlWriteCipher.AES_256, ExcelOoxmlWriteHash.SHA_512);
  }

  private static ExcelOoxmlWriteCipher normalizeCipher(ExcelOoxmlWriteCipher cipher) {
    return cipher == null ? ExcelOoxmlWriteCipher.AES_256 : cipher;
  }

  private static ExcelOoxmlWriteHash normalizeHash(ExcelOoxmlWriteHash hash) {
    return hash == null ? ExcelOoxmlWriteHash.SHA_512 : hash;
  }
}
