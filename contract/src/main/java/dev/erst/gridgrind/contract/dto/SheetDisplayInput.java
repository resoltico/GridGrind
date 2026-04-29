package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

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
    Objects.requireNonNull(displayGridlines, "displayGridlines must not be null");
    Objects.requireNonNull(displayZeros, "displayZeros must not be null");
    Objects.requireNonNull(displayRowColHeadings, "displayRowColHeadings must not be null");
    Objects.requireNonNull(displayFormulas, "displayFormulas must not be null");
    Objects.requireNonNull(rightToLeft, "rightToLeft must not be null");
  }

  @JsonCreator
  static SheetDisplayInput create(
      @JsonProperty("displayGridlines") Boolean displayGridlines,
      @JsonProperty("displayZeros") Boolean displayZeros,
      @JsonProperty("displayRowColHeadings") Boolean displayRowColHeadings,
      @JsonProperty("displayFormulas") Boolean displayFormulas,
      @JsonProperty("rightToLeft") Boolean rightToLeft) {
    SheetDisplayInput defaults = defaults();
    return new SheetDisplayInput(
        displayGridlines == null ? defaults.displayGridlines() : displayGridlines,
        displayZeros == null ? defaults.displayZeros() : displayZeros,
        displayRowColHeadings == null ? defaults.displayRowColHeadings() : displayRowColHeadings,
        displayFormulas == null ? defaults.displayFormulas() : displayFormulas,
        rightToLeft == null ? defaults.rightToLeft() : rightToLeft);
  }
}
