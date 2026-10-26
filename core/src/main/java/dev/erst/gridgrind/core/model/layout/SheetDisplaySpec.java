package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Screen-facing sheet display flags authored as part of sheet-presentation state. */
@CatalogType(
    id = "SheetDisplaySpec",
    summary =
        "Screen-facing sheet display flags authored as part of sheet-presentation state. Supply all display flags explicitly.",
    displayOrder = 67)
public record SheetDisplaySpec(
    boolean displayGridlines,
    boolean displayZeros,
    boolean displayRowColHeadings,
    boolean displayFormulas,
    boolean rightToLeft) {
  /** Returns the effective Excel defaults for one unconfigured worksheet view. */
  public static SheetDisplaySpec defaults() {
    return new SheetDisplaySpec(true, true, true, false, false);
  }

  public SheetDisplaySpec {}
}
