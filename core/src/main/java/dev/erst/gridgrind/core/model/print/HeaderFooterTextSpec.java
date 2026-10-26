package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.Objects;

/** Plain left, center, and right header or footer text segments in protocol form. */
@CatalogType(
    id = "HeaderFooterTextSpec",
    summary =
        "Plain left, center, and right header or footer text segments. Supply all three fields explicitly.",
    displayOrder = 59)
public record HeaderFooterTextSpec(
    TextSourceSpec left, TextSourceSpec center, TextSourceSpec right) {
  /** Returns the default blank header or footer text payload. */
  public static HeaderFooterTextSpec blank() {
    return new HeaderFooterTextSpec(defaultSegment(), defaultSegment(), defaultSegment());
  }

  public HeaderFooterTextSpec {
    Objects.requireNonNull(left, "left must not be null");
    Objects.requireNonNull(center, "center must not be null");
    Objects.requireNonNull(right, "right must not be null");
  }

  private static TextSourceSpec defaultSegment() {
    return new TextSourceSpec.Inline("");
  }
}
