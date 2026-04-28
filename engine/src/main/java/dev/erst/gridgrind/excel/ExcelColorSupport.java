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
        switch (color) {
          case ExcelColor.Rgb rgb -> ExcelRgbColorSupport.toXssfColor(workbook, rgb.rgb());
          case ExcelColor.Theme _ -> new XSSFColor(workbook.getStylesSource().getIndexedColors());
          case ExcelColor.Indexed _ -> new XSSFColor(workbook.getStylesSource().getIndexedColors());
        };
    switch (color) {
      case ExcelColor.Rgb rgb -> applyTint(xssfColor, rgb.tint());
      case ExcelColor.Theme theme -> {
        xssfColor.setTheme(theme.theme());
        applyTint(xssfColor, theme.tint());
      }
      case ExcelColor.Indexed indexed -> {
        xssfColor.setIndexed(indexed.indexed());
        applyTint(xssfColor, indexed.tint());
      }
    }
    return xssfColor;
  }

  static ExcelColor copyOf(ExcelColorSnapshot color) {
    return color == null
        ? null
        : switch (color) {
          case ExcelColorSnapshot.Rgb rgb -> ExcelColor.rgb(rgb.rgb(), rgb.tint());
          case ExcelColorSnapshot.Theme theme -> ExcelColor.theme(theme.theme(), theme.tint());
          case ExcelColorSnapshot.Indexed indexed ->
              ExcelColor.indexed(indexed.indexed(), indexed.tint());
        };
  }

  private static void applyTint(XSSFColor color, Double tint) {
    if (tint != null) {
      color.setTint(tint);
    }
  }
}
