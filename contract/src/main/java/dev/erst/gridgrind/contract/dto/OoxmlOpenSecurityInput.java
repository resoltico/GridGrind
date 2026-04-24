package dev.erst.gridgrind.contract.dto;

/** Optional OOXML package-open settings for encrypted existing workbook sources. */
public record OoxmlOpenSecurityInput(String password) {
  public OoxmlOpenSecurityInput {
    password = normalizeOptional(password, "password").orElse(null);
  }

  boolean isEmpty() {
    return password == null;
  }

  private static java.util.Optional<String> normalizeOptional(String value, String fieldName) {
    if (value == null) {
      return java.util.Optional.empty();
    }
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return java.util.Optional.of(value);
  }
}
