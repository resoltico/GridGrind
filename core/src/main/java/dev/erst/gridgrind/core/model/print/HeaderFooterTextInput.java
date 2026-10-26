package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.TextSourceInput;
import java.util.Objects;

/** Plain left, center, and right header or footer text segments in protocol form. */
@CatalogType(
    id = "HeaderFooterTextInput",
    summary =
        "Plain left, center, and right header or footer text segments. Supply all three fields explicitly.",
    displayOrder = 59)
public record HeaderFooterTextInput(
    TextSourceInput left, TextSourceInput center, TextSourceInput right) {
  /** Returns the default blank header or footer text payload. */
  public static HeaderFooterTextInput blank() {
    return new HeaderFooterTextInput(defaultSegment(), defaultSegment(), defaultSegment());
  }

  public HeaderFooterTextInput {
    Objects.requireNonNull(left, "left must not be null");
    Objects.requireNonNull(center, "center must not be null");
    Objects.requireNonNull(right, "right must not be null");
  }

  private static TextSourceInput defaultSegment() {
    return new TextSourceInput.Inline("");
  }
}
