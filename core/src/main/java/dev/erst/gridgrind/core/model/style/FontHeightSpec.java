package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.math.BigDecimal;
import java.util.Objects;

/** JSON-friendly typed font height input used by style patches. */
@Union(discriminator = "type", catalog = true, displayOrder = 54)
public sealed interface FontHeightSpec {

  /** Returns the exact height in whole twips for workbook authoring. */
  default int twips() {
    return switch (this) {
      case Points height -> height.points().multiply(BigDecimal.valueOf(20)).intValueExact();
      case Twips height -> height.twips();
    };
  }

  /** Returns the height in point units for workbook authoring. */
  default BigDecimal points() {
    return switch (this) {
      case Points height -> height.points();
      case Twips height -> BigDecimal.valueOf(height.twips()).divide(BigDecimal.valueOf(20));
    };
  }

  /** Font height expressed in point units, such as {@code 11} or {@code 11.5}. */
  @Variant(
      id = "POINTS",
      summary =
          "Specify font height in point units. Write format: {\"type\":\"POINTS\",\"points\":13}. Read-back (GET_CELLS, GET_WINDOW): style.fontHeight is {\"twips\":260,\"points\":13} with both fields present, not this discriminated type format.")
  record Points(BigDecimal points) implements FontHeightSpec {
    public Points {
      validatePoints(points);
    }
  }

  /** Font height expressed in exact twips, where one point equals twenty twips. */
  @Variant(
      id = "TWIPS",
      summary =
          "Specify font height in exact twips (20 twips = 1 point). Write format: {\"type\":\"TWIPS\",\"twips\":260}. Read-back returns the same plain object shape as POINTS.")
  record Twips(int twips) implements FontHeightSpec {
    public Twips {
      validateTwips(twips);
    }
  }

  private static void validatePoints(BigDecimal points) {
    Objects.requireNonNull(points, "points must not be null");
    if (points.signum() <= 0) {
      throw new IllegalArgumentException("points must be greater than 0");
    }
    try {
      validateTwips(points.multiply(BigDecimal.valueOf(20)).intValueExact());
    } catch (ArithmeticException exception) {
      throw new IllegalArgumentException(
          "points must resolve exactly to whole twips (1/20 points)", exception);
    }
  }

  private static void validateTwips(int twips) {
    if (twips <= 0 || twips > Short.MAX_VALUE) {
      throw new IllegalArgumentException(
          "twips must be between 1 and " + Short.MAX_VALUE + " (inclusive)");
    }
  }
}
