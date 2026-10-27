package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import java.util.Objects;

/** OOXML package-encryption settings applied during workbook persistence. */
public record ResolvedEncryption(
    String password, ExcelOoxmlWriteCipher cipher, ExcelOoxmlWriteHash hash) {
  public ResolvedEncryption {
    password = normalizeRequired(password, "password");
    cipher = Objects.requireNonNullElse(cipher, ExcelOoxmlWriteCipher.AES_256);
    hash = Objects.requireNonNullElse(hash, ExcelOoxmlWriteHash.SHA_512);
  }

  private static String normalizeRequired(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  @Override
  public String toString() {
    return "ResolvedEncryption[password=redacted, cipher=" + cipher + ", hash=" + hash + "]";
  }
}
