package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Screen-facing sheet display flags authored as part of sheet-presentation state. */
@CatalogType(
    id = "SheetDisplayInput",
    summary =
        "Screen-facing sheet display flags authored as part of sheet-presentation state. Supply all display flags explicitly.",
    displayOrder = 67)
public record SheetDisplayInput(
    boolean displayGridlines,
    boolean displayZeros,
    boolean displayRowColHeadings,
    boolean displayFormulas,
    boolean rightToLeft) {
  /** Returns the effective Excel defaults for one unconfigured worksheet view. */
  public static SheetDisplayInput defaults() {
    return new SheetDisplayInput(true, true, true, false, false);
  }

  public SheetDisplayInput {}
}
