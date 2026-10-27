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
    id = "OoxmlEncryptionSpec",
    summary =
        "OOXML package-encryption settings for workbook persistence. GridGrind writes AGILE packages only. cipher defaults to AES_256 and hash defaults to SHA_512 when omitted. Supported ciphers: AES_256 and AES_192. Supported hashes: SHA_512, SHA_384, and SHA_256 ("
            + dev.erst.gridgrind.core.limits.Limit
                .ID_OOXML_WRITE_ENCRYPTION_IS_AGILE_ONLY_AND_STRONG_ONLY
            + ").",
    displayOrder = 21)
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_OOXML_WRITE_ENCRYPTION_IS_AGILE_ONLY_AND_STRONG_ONLY
})
public record OoxmlEncryptionSpec(
    SecretReference passwordRef,
    @ProtocolField(optional = true) @WireDefault("AES_256") ExcelOoxmlWriteCipher cipher,
    @ProtocolField(optional = true) @WireDefault("SHA_512") ExcelOoxmlWriteHash hash) {

  public OoxmlEncryptionSpec {
    Objects.requireNonNull(passwordRef, "passwordRef must not be null");
    Objects.requireNonNull(cipher, "cipher must not be null");
    Objects.requireNonNull(hash, "hash must not be null");
  }

  /** Creates one strong OOXML encryption payload with GridGrind's default write envelope. */
  public static OoxmlEncryptionSpec strong(SecretReference passwordRef) {
    return new OoxmlEncryptionSpec(
        passwordRef, ExcelOoxmlWriteCipher.AES_256, ExcelOoxmlWriteHash.SHA_512);
  }
}
