package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

/** Converts POI workbook colors into factual snapshot structures without flattening semantics. */
final class ExcelColorSnapshotSupport {
  private ExcelColorSnapshotSupport() {}

  /** Returns one factual snapshot for the supplied workbook color, or empty when absent. */
  static Optional<ExcelColorSnapshot> snapshot(XSSFColor color) {
    if (color == null) {
      return Optional.empty();
    }
    String rgb = color.isRGB() ? ExcelRgbColorSupport.toRgbHex(color).orElse(null) : null;
    Integer theme = color.isThemed() ? color.getTheme() : null;
    Integer indexed = color.isIndexed() ? Short.toUnsignedInt(color.getIndexed()) : null;
    Optional<Double> tint = color.hasTint() ? Optional.of(color.getTint()) : Optional.empty();
    if (rgb != null) {
      return Optional.of(ExcelColorSnapshot.rgb(rgb, tint));
    }
    if (theme != null) {
      return Optional.of(ExcelColorSnapshot.theme(theme, tint));
    }
    if (indexed != null) {
      return Optional.of(ExcelColorSnapshot.indexed(indexed, tint));
    }
    throw new IllegalStateException(
        "Workbook color could not be normalized: no RGB, theme, or indexed payload was present"
            + (color.isAuto() ? " (automatic color)." : "."));
  }

  /** Returns one factual snapshot for the supplied raw workbook color XML, or empty when absent. */
  static Optional<ExcelColorSnapshot> snapshot(XSSFWorkbook workbook, CTColor color) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    if (color == null) {
      return Optional.empty();
    }
    XSSFColor xssfColor = XSSFColor.from(color, workbook.getStylesSource().getIndexedColors());
    ThemesTable themes = workbook.getStylesSource().getTheme();
    if (themes != null) {
      themes.inheritFromThemeAsRequired(xssfColor);
    }
    return snapshot(xssfColor);
  }
}
