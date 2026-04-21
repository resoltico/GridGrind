package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Explicit OOXML package-open security alternatives for plain and encrypted workbook sources. */
public sealed interface ExcelOoxmlOpenOptions
    permits ExcelOoxmlOpenOptions.Unencrypted, ExcelOoxmlOpenOptions.Encrypted {

  /** Opens a plain `.xlsx` workbook with no decryption password. */
  record Unencrypted() implements ExcelOoxmlOpenOptions {}

  /** Opens an encrypted OOXML workbook using the supplied verified password candidate. */
  record Encrypted(String password) implements ExcelOoxmlOpenOptions {
    public Encrypted {
      Objects.requireNonNull(password, "password must not be null");
      if (password.isBlank()) {
        throw new IllegalArgumentException("password must not be blank");
      }
    }
  }
}
