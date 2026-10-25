package dev.erst.gridgrind.contract.dto;

/** Factual screen-facing sheet display flags reported for one worksheet view. */
public record SheetDisplayReport(
    boolean displayGridlines,
    boolean displayZeros,
    boolean displayRowColHeadings,
    boolean displayFormulas,
    boolean rightToLeft) {
  /** Returns the effective Excel defaults for one unconfigured worksheet view. */
  public static SheetDisplayReport defaults() {
    return new SheetDisplayReport(true, true, true, false, false);
  }
}
