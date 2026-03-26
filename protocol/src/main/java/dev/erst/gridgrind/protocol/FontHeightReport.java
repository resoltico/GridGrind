package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelFontHeight;
import java.math.BigDecimal;
import java.util.Objects;

/** Exact and human-friendly font height facts reported during workbook analysis. */
public record FontHeightReport(int twips, BigDecimal points) {
  public FontHeightReport {
    Objects.requireNonNull(points, "points must not be null");
    ExcelFontHeight fontHeight = new ExcelFontHeight(twips);
    if (fontHeight.points().compareTo(points.stripTrailingZeros()) != 0) {
      throw new IllegalArgumentException("points must match twips exactly");
    }
    points = points.stripTrailingZeros();
  }

  /** Creates a protocol report from the canonical engine font height value. */
  public static FontHeightReport fromExcelFontHeight(ExcelFontHeight fontHeight) {
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
    return new FontHeightReport(fontHeight.twips(), fontHeight.points());
  }
}
