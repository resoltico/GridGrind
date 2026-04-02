package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Plain left, center, and right header or footer text segments. */
public record ExcelHeaderFooterText(String left, String center, String right) {
  /** Returns the default blank header or footer text. */
  public static ExcelHeaderFooterText blank() {
    return new ExcelHeaderFooterText("", "", "");
  }

  public ExcelHeaderFooterText {
    left = requireNonNull(left, "left");
    center = requireNonNull(center, "center");
    right = requireNonNull(right, "right");
  }

  /** Returns whether every segment is blank. */
  public boolean isBlank() {
    return left.isEmpty() && center.isEmpty() && right.isEmpty();
  }

  private static String requireNonNull(String value, String fieldName) {
    return Objects.requireNonNull(value, fieldName + " must not be null");
  }
}
