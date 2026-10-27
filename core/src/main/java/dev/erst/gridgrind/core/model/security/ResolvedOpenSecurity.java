package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Explicit OOXML package-open security alternatives for plain and encrypted workbook sources. */
public sealed interface ResolvedOpenSecurity
    permits ResolvedOpenSecurity.Unencrypted, ResolvedOpenSecurity.Encrypted {

  /** Opens a plain `.xlsx` workbook with no decryption password. */
  record Unencrypted() implements ResolvedOpenSecurity {}

  /** Opens an encrypted OOXML workbook using the supplied verified password candidate. */
  record Encrypted(String password) implements ResolvedOpenSecurity {
    public Encrypted {
      password = TextValidation.requireNonBlank(password, "password");
    }

    @Override
    public String toString() {
      return "Encrypted[password=redacted]";
    }
  }
}
