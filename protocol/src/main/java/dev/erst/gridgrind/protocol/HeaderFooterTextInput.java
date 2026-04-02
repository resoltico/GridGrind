package dev.erst.gridgrind.protocol;

import java.util.Objects;

/** Plain left, center, and right header or footer text segments in protocol form. */
public record HeaderFooterTextInput(String left, String center, String right) {
  /** Returns the default blank header or footer text payload. */
  public static HeaderFooterTextInput blank() {
    return new HeaderFooterTextInput("", "", "");
  }

  public HeaderFooterTextInput {
    left = Objects.requireNonNullElse(left, "");
    center = Objects.requireNonNullElse(center, "");
    right = Objects.requireNonNullElse(right, "");
  }
}
