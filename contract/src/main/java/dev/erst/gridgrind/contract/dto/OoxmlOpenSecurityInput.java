package dev.erst.gridgrind.contract.dto;

/** Optional OOXML package-open settings for encrypted existing workbook sources. */
public record OoxmlOpenSecurityInput(String password) {
  public OoxmlOpenSecurityInput {
    password = normalizeOptional(password, "password");
  }

  boolean isEmpty() {
    return password == null;
  }

  private static String normalizeOptional(String value, String fieldName) {
    if (value == null) {
      return null;
    }
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
