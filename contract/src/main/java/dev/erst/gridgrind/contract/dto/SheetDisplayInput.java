package dev.erst.gridgrind.contract.dto;

/** Screen-facing sheet display flags authored as part of sheet-presentation state. */
public record SheetDisplayInput(
    Boolean displayGridlines,
    Boolean displayZeros,
    Boolean displayRowColHeadings,
    Boolean displayFormulas,
    Boolean rightToLeft) {
  /** Returns the effective Excel defaults for one unconfigured worksheet view. */
  public static SheetDisplayInput defaults() {
    return new SheetDisplayInput(true, true, true, false, false);
  }

  public SheetDisplayInput {
    displayGridlines = displayGridlines == null ? defaults().displayGridlines() : displayGridlines;
    displayZeros = displayZeros == null ? defaults().displayZeros() : displayZeros;
    displayRowColHeadings =
        displayRowColHeadings == null ? defaults().displayRowColHeadings() : displayRowColHeadings;
    displayFormulas = displayFormulas == null ? defaults().displayFormulas() : displayFormulas;
    rightToLeft = rightToLeft == null ? defaults().rightToLeft() : rightToLeft;
  }
}
