package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;

/** Screen-facing sheet display flags authored as part of sheet-presentation state. */
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

  /** Creates sheet-display settings while applying the documented worksheet-view defaults. */
  @JsonCreator
  public SheetDisplayInput(
      Boolean displayGridlines,
      Boolean displayZeros,
      Boolean displayRowColHeadings,
      Boolean displayFormulas,
      Boolean rightToLeft) {
    this(
        java.util.Objects.requireNonNull(displayGridlines, "displayGridlines must not be null")
            .booleanValue(),
        java.util.Objects.requireNonNull(displayZeros, "displayZeros must not be null")
            .booleanValue(),
        java.util.Objects.requireNonNull(
                displayRowColHeadings, "displayRowColHeadings must not be null")
            .booleanValue(),
        java.util.Objects.requireNonNull(displayFormulas, "displayFormulas must not be null")
            .booleanValue(),
        java.util.Objects.requireNonNull(rightToLeft, "rightToLeft must not be null")
            .booleanValue());
  }
}
