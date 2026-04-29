package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

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
    Objects.requireNonNull(left, "left must not be null");
    Objects.requireNonNull(center, "center must not be null");
    Objects.requireNonNull(right, "right must not be null");
  }

  @JsonCreator
  static HeaderFooterTextInput create(
      @JsonProperty("left") TextSourceInput left,
      @JsonProperty("center") TextSourceInput center,
      @JsonProperty("right") TextSourceInput right) {
    HeaderFooterTextInput defaults = blank();
    return new HeaderFooterTextInput(
        left == null ? defaults.left() : left,
        center == null ? defaults.center() : center,
        right == null ? defaults.right() : right);
  }
}
