package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

/** Converts POI workbook colors into factual snapshot structures without flattening semantics. */
final class ExcelColorSnapshotSupport {
  private ExcelColorSnapshotSupport() {}

  /** Returns one factual snapshot for the supplied workbook color, or null when absent. */
  static ExcelColorSnapshot snapshot(XSSFColor color) {
    if (color == null) {
      return java.util.Optional.<ExcelColorSnapshot>empty().orElse(null);
    }
    String rgb = color.isRGB() ? ExcelRgbColorSupport.toRgbHex(color).orElse(null) : null;
    Integer theme = color.isThemed() ? color.getTheme() : null;
    Integer indexed = color.isIndexed() ? Short.toUnsignedInt(color.getIndexed()) : null;
    Double tint = color.hasTint() ? color.getTint() : null;
    return rgb != null
        ? ExcelColorSnapshot.rgb(rgb, tint)
        : theme != null
            ? ExcelColorSnapshot.theme(theme, tint)
            : indexed != null
                ? ExcelColorSnapshot.indexed(indexed, tint)
                : java.util.Optional.<ExcelColorSnapshot>empty().orElse(null);
  }

  /** Returns one factual snapshot for the supplied raw workbook color XML, or null when absent. */
  static ExcelColorSnapshot snapshot(XSSFWorkbook workbook, CTColor color) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    if (color == null) {
      return java.util.Optional.<ExcelColorSnapshot>empty().orElse(null);
    }
    XSSFColor xssfColor = XSSFColor.from(color, workbook.getStylesSource().getIndexedColors());
    ThemesTable themes = workbook.getStylesSource().getTheme();
    if (themes != null) {
      themes.inheritFromThemeAsRequired(xssfColor);
    }
    return snapshot(xssfColor);
  }
}
