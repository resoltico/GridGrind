package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import java.util.Objects;

/** OOXML package-encryption settings applied during workbook persistence. */
public record OoxmlEncryptionInput(
    @ProtocolField(secret = true) String password,
    @ProtocolField(optional = true) ExcelOoxmlWriteCipher cipher,
    @ProtocolField(optional = true) ExcelOoxmlWriteHash hash) {
  /** Reads one OOXML encryption block while applying the documented omission default. */
  @JsonCreator
  static OoxmlEncryptionInput create(
      @JsonProperty("password") String password,
      @JsonProperty("cipher") ExcelOoxmlWriteCipher cipher,
      @JsonProperty("hash") ExcelOoxmlWriteHash hash) { // LIM-038
    return new OoxmlEncryptionInput(password, normalizeCipher(cipher), normalizeHash(hash));
  }

  public OoxmlEncryptionInput {
    password = normalizeRequired(password, "password");
    Objects.requireNonNull(cipher, "cipher must not be null");
    Objects.requireNonNull(hash, "hash must not be null");
  }

  /** Creates one strong OOXML encryption payload with GridGrind's default write envelope. */
  public static OoxmlEncryptionInput strong(String password) {
    return new OoxmlEncryptionInput(
        password, ExcelOoxmlWriteCipher.AES_256, ExcelOoxmlWriteHash.SHA_512);
  }

  private static String normalizeRequired(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static ExcelOoxmlWriteCipher normalizeCipher(ExcelOoxmlWriteCipher cipher) {
    return cipher == null ? ExcelOoxmlWriteCipher.AES_256 : cipher;
  }

  private static ExcelOoxmlWriteHash normalizeHash(ExcelOoxmlWriteHash hash) {
    return hash == null ? ExcelOoxmlWriteHash.SHA_512 : hash;
  }
}
