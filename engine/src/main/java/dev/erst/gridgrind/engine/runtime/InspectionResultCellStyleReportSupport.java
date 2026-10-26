package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Converts workbook style snapshots into protocol style report records. */
final class InspectionResultCellStyleReportSupport {
  private InspectionResultCellStyleReportSupport() {}

  static CellStyleReport toCellStyleReport(ExcelCellStyleSnapshot style) {
    return new CellStyleReport(
        style.numberFormat(),
        style.alignment(),
        toCellFontReport(style.font()),
        style.fill(),
        new CellBorderReport(
            toCellBorderSideReport(style.border().top()),
            toCellBorderSideReport(style.border().right()),
            toCellBorderSideReport(style.border().bottom()),
            toCellBorderSideReport(style.border().left())),
        style.protection());
  }

  static Optional<CellColorReport> toCellColorReport(@Nullable ExcelColor color) {
    return color == null
        ? Optional.empty()
        : Optional.of(
            switch (color) {
              case ExcelColor.Rgb rgb ->
                  rgb.tint().isPresent()
                      ? CellColorReport.rgb(new RgbHex(rgb.rgb()), rgb.tint().orElseThrow())
                      : CellColorReport.rgb(new RgbHex(rgb.rgb()));
              case ExcelColor.Theme theme ->
                  theme.tint().isPresent()
                      ? CellColorReport.theme(theme.theme(), theme.tint().orElseThrow())
                      : CellColorReport.theme(theme.theme());
              case ExcelColor.Indexed indexed ->
                  indexed.tint().isPresent()
                      ? CellColorReport.indexed(indexed.indexed(), indexed.tint().orElseThrow())
                      : CellColorReport.indexed(indexed.indexed());
            });
  }

  static FontHeightReport toFontHeightReport(ExcelFontHeight fontHeight) {
    return new FontHeightReport(fontHeight.twips(), fontHeight.points());
  }

  static CellFontReport toCellFontReport(ExcelCellFontSnapshot font) {
    return new CellFontReport(
        font.bold(),
        font.italic(),
        font.fontName(),
        toFontHeightReport(font.fontHeight()),
        font.fontColor(),
        font.underline(),
        font.strikeout());
  }

  static CellBorderSideReport toCellBorderSideReport(ExcelBorderSideSnapshot side) {
    if (side.style() == ExcelBorderStyle.NONE) {
      return new CellBorderSideReport.None();
    }
    return Optional.ofNullable(side.color())
        .<CellBorderSideReport>map(color -> new CellBorderSideReport.Colored(side.style(), color))
        .orElseGet(() -> new CellBorderSideReport.DefaultColor(side.style()));
  }
}
