package dev.erst.gridgrind.core.model.security;

import java.util.Objects;

/** Explicit OOXML package-open security alternatives for plain and encrypted workbook sources. */
public sealed interface ResolvedOpenSecurity
    permits ResolvedOpenSecurity.Unencrypted, ResolvedOpenSecurity.Encrypted {

  /** Opens a plain `.xlsx` workbook with no decryption password. */
  record Unencrypted() implements ResolvedOpenSecurity {}

  /** Opens an encrypted OOXML workbook using the supplied verified password candidate. */
  record Encrypted(String password) implements ResolvedOpenSecurity {
    public Encrypted {
      Objects.requireNonNull(password, "password must not be null");
      if (password.isBlank()) {
        throw new IllegalArgumentException("password must not be blank");
      }
    }

    @Override
    public String toString() {
      return "Encrypted[password=redacted]";
    }
  }
}
