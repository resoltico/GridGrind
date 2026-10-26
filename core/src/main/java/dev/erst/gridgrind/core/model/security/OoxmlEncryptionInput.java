package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
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
    @ProtocolField(optional = true) @WireDefault("AES_256") ExcelOoxmlWriteCipher cipher,
    @ProtocolField(optional = true) @WireDefault("SHA_512") ExcelOoxmlWriteHash hash) {

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
}
