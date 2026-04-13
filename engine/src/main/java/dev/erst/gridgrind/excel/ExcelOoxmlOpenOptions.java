package dev.erst.gridgrind.excel;

/** Optional OOXML package-open security settings for encrypted workbook sources. */
public record ExcelOoxmlOpenOptions(String password) {
  public ExcelOoxmlOpenOptions {
    password = normalizeOptional(password, "password");
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
