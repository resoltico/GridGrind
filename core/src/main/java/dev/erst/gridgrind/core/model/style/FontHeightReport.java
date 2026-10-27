package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.math.BigDecimal;
import java.util.Objects;

/** Exact and human-friendly font height facts reported during workbook analysis. */
@CatalogType(
    id = "FontHeightReport",
    summary = "Exact font-height report expressed in twips and points.",
    displayOrder = 83)
public record FontHeightReport(int twips, BigDecimal points) {
  /** Derives the exact point value from a persisted twip height. */
  public static FontHeightReport fromTwips(int twips) {
    return new FontHeightReport(twips, BigDecimal.valueOf(twips).divide(BigDecimal.valueOf(20)));
  }

  /** Converts point units exactly, rejecting fractional twips. */
  public static FontHeightReport fromPoints(BigDecimal points) {
    Objects.requireNonNull(points, "points must not be null");
    if (points.signum() <= 0) {
      throw new IllegalArgumentException("points must be greater than 0");
    }
    try {
      return fromTwips(points.multiply(BigDecimal.valueOf(20)).intValueExact());
    } catch (ArithmeticException exception) {
      throw new IllegalArgumentException(
          "points must resolve exactly to whole twips (1/20 points)", exception);
    }
  }

  public FontHeightReport {
    Objects.requireNonNull(points, "points must not be null");
    if (twips <= 0 || twips > Short.MAX_VALUE) {
      throw new IllegalArgumentException(
          "twips must be between 1 and " + Short.MAX_VALUE + " (inclusive)");
    }
    BigDecimal normalizedPoints = BigDecimal.valueOf(twips).divide(BigDecimal.valueOf(20));
    if (normalizedPoints.compareTo(points.stripTrailingZeros()) != 0) {
      throw new IllegalArgumentException("points must match twips exactly");
    }
    points = points.stripTrailingZeros();
  }
}
