package dev.erst.gridgrind.contract.dto;

import java.math.BigDecimal;
import java.util.Objects;

/** Exact and human-friendly font height facts reported during workbook analysis. */
public record FontHeightReport(int twips, BigDecimal points) {
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
