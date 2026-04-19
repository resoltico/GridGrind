package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;

/** Plain left, center, and right header or footer text segments in protocol form. */
public record HeaderFooterTextInput(
    TextSourceInput left, TextSourceInput center, TextSourceInput right) {
  /** Returns the default blank header or footer text payload. */
  public static HeaderFooterTextInput blank() {
    return new HeaderFooterTextInput(
        new TextSourceInput.Inline(""),
        new TextSourceInput.Inline(""),
        new TextSourceInput.Inline(""));
  }

  public HeaderFooterTextInput {
    left = left == null ? new TextSourceInput.Inline("") : left;
    center = center == null ? new TextSourceInput.Inline("") : center;
    right = right == null ? new TextSourceInput.Inline("") : right;
  }
}
