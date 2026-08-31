package dev.erst.gridgrind.contract.source;

import java.util.Objects;

/** Validates authored text before it can be persisted as OOXML character data. */
public final class OoxmlTextValidation {
  private OoxmlTextValidation() {}

  /** Rejects every Unicode code point forbidden by XML 1.0 character-data rules. */
  public static String requireXml10CharacterData(String text, String fieldName) {
    Objects.requireNonNull(text, fieldName + " must not be null");
    Objects.requireNonNull(fieldName, "fieldName must not be null");
    int index = 0;
    while (index < text.length()) {
      int codePoint = text.codePointAt(index);
      if (!isXml10Character(codePoint)) {
        throw new IllegalArgumentException(
            fieldName
                + " contains XML 1.0-forbidden code point U+"
                + String.format("%04X", codePoint)
                + " at code-point index "
                + text.codePointCount(0, index));
      }
      index += Character.charCount(codePoint);
    }
    return text;
  }

  private static boolean isXml10Character(int codePoint) {
    return codePoint == 0x9
        || codePoint == 0xA
        || codePoint == 0xD
        || (codePoint >= 0x20 && codePoint <= 0xD7FF)
        || (codePoint >= 0xE000 && codePoint <= 0xFFFD)
        || codePoint >= 0x10000;
  }
}
