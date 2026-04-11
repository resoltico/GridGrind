package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Shared mutable-workbook color conversion helpers. */
final class ExcelColorSupport {
  private ExcelColorSupport() {}

  /** Converts one authored color into the matching POI workbook color. */
  static XSSFColor toXssfColor(XSSFWorkbook workbook, ExcelColor color) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(color, "color must not be null");

    XSSFColor xssfColor =
        color.rgb() == null
            ? new XSSFColor(workbook.getStylesSource().getIndexedColors())
            : ExcelRgbColorSupport.toXssfColor(workbook, color.rgb());
    if (color.theme() != null) {
      xssfColor.setTheme(color.theme());
    }
    if (color.indexed() != null) {
      xssfColor.setIndexed(color.indexed());
    }
    if (color.tint() != null) {
      xssfColor.setTint(color.tint());
    }
    return xssfColor;
  }
}
