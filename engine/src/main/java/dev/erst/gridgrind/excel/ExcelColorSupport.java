package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;

/** Shared mutable-workbook color conversion helpers. */
final class ExcelColorSupport {
  private ExcelColorSupport() {}

  /** Converts one authored color into the matching POI workbook color. */
  static XSSFColor toXssfColor(XSSFWorkbook workbook, ColorSpec color) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(color, "color must not be null");

    XSSFColor xssfColor =
        switch (color) {
          case ColorSpec.Rgb rgb -> ExcelRgbColorSupport.toXssfColor(workbook, rgb.rgb().value());
          case ColorSpec.Theme _ -> new XSSFColor(workbook.getStylesSource().getIndexedColors());
          case ColorSpec.Indexed _ -> new XSSFColor(workbook.getStylesSource().getIndexedColors());
        };
    switch (color) {
      case ColorSpec.Rgb rgb -> applyTint(xssfColor, rgb.tint());
      case ColorSpec.Theme theme -> {
        xssfColor.setTheme(theme.theme());
        applyTint(xssfColor, theme.tint());
      }
      case ColorSpec.Indexed indexed -> {
        xssfColor.setIndexed(indexed.indexed());
        applyTint(xssfColor, indexed.tint());
      }
    }
    return xssfColor;
  }

  static @Nullable ColorSpec copyOf(@Nullable CellColorReport color) {
    return color == null
        ? null
        : switch (color) {
          case CellColorReport.Rgb rgb -> new ColorSpec.Rgb(rgb.rgb(), rgb.tint());
          case CellColorReport.Theme theme -> new ColorSpec.Theme(theme.theme(), theme.tint());
          case CellColorReport.Indexed indexed ->
              new ColorSpec.Indexed(indexed.indexed(), indexed.tint());
        };
  }

  private static void applyTint(XSSFColor color, java.util.Optional<Double> tint) {
    tint.ifPresent(color::setTint);
  }
}
