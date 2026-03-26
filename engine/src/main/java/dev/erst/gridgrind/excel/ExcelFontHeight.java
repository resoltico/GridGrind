package dev.erst.gridgrind.excel;

import java.math.BigDecimal;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Font;

/** Exact Excel font height stored in twips, where one point equals twenty twips. */
public record ExcelFontHeight(int twips) {
  public ExcelFontHeight {
    if (twips <= 0 || twips > Short.MAX_VALUE) {
      throw new IllegalArgumentException(
          "twips must be between 1 and " + Short.MAX_VALUE + " (inclusive)");
    }
  }

  /** Creates a font height from point units, requiring exact conversion to whole twips. */
  public static ExcelFontHeight fromPoints(BigDecimal points) {
    Objects.requireNonNull(points, "points must not be null");
    if (points.signum() <= 0) {
      throw new IllegalArgumentException("points must be greater than 0");
    }
    BigDecimal twips = points.multiply(BigDecimal.valueOf(Font.TWIPS_PER_POINT));
    try {
      return new ExcelFontHeight(twips.intValueExact());
    } catch (ArithmeticException exception) {
      throw new IllegalArgumentException(
          "points must resolve exactly to whole twips (1/20 points)", exception);
    }
  }

  /** Returns the font height in point units. */
  public BigDecimal points() {
    return BigDecimal.valueOf(twips)
        .divide(BigDecimal.valueOf(Font.TWIPS_PER_POINT))
        .stripTrailingZeros();
  }
}
