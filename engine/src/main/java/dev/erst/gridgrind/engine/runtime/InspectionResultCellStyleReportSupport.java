package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelGradientFillSnapshot;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Converts workbook style snapshots into protocol style report records. */
final class InspectionResultCellStyleReportSupport {
  private InspectionResultCellStyleReportSupport() {}

  static CellStyleReport toCellStyleReport(ExcelCellStyleSnapshot style) {
    return new CellStyleReport(
        style.numberFormat(),
        new CellAlignmentReport(
            style.alignment().wrapText(),
            style.alignment().horizontalAlignment(),
            style.alignment().verticalAlignment(),
            style.alignment().textRotation(),
            style.alignment().indentation()),
        toCellFontReport(style.font()),
        toCellFillReport(style.fill()),
        new CellBorderReport(
            toCellBorderSideReport(style.border().top()),
            toCellBorderSideReport(style.border().right()),
            toCellBorderSideReport(style.border().bottom()),
            toCellBorderSideReport(style.border().left())),
        new CellProtectionReport(style.protection().locked(), style.protection().hiddenFormula()));
  }

  static Optional<CellColorReport> toCellColorReport(@Nullable ExcelColorSnapshot color) {
    return color == null
        ? Optional.empty()
        : Optional.of(
            switch (color) {
              case ExcelColorSnapshot.Rgb rgb ->
                  rgb.tint().isPresent()
                      ? CellColorReport.rgb(rgb.rgb(), rgb.tint().orElseThrow())
                      : CellColorReport.rgb(rgb.rgb());
              case ExcelColorSnapshot.Theme theme ->
                  theme.tint().isPresent()
                      ? CellColorReport.theme(theme.theme(), theme.tint().orElseThrow())
                      : CellColorReport.theme(theme.theme());
              case ExcelColorSnapshot.Indexed indexed ->
                  indexed.tint().isPresent()
                      ? CellColorReport.indexed(indexed.indexed(), indexed.tint().orElseThrow())
                      : CellColorReport.indexed(indexed.indexed());
            });
  }

  static Optional<CellColorReport> toCellColorReport(@Nullable ExcelColor color) {
    return color == null
        ? Optional.empty()
        : Optional.of(
            switch (color) {
              case ExcelColor.Rgb rgb ->
                  rgb.tint().isPresent()
                      ? CellColorReport.rgb(rgb.rgb(), rgb.tint().orElseThrow())
                      : CellColorReport.rgb(rgb.rgb());
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
        toCellColorReport(font.fontColor()).orElse(null),
        font.underline(),
        font.strikeout());
  }

  static CellBorderSideReport toCellBorderSideReport(ExcelBorderSideSnapshot side) {
    if (side.style() == ExcelBorderStyle.NONE) {
      return new CellBorderSideReport.None();
    }
    return toCellColorReport(side.color())
        .<CellBorderSideReport>map(color -> new CellBorderSideReport.Colored(side.style(), color))
        .orElseGet(() -> new CellBorderSideReport.DefaultColor(side.style()));
  }

  private static CellFillReport toCellFillReport(ExcelCellFillSnapshot fill) {
    return switch (fill) {
      case ExcelCellFillSnapshot.PatternOnly pattern -> CellFillReport.pattern(pattern.pattern());
      case ExcelCellFillSnapshot.PatternForeground pattern ->
          CellFillReport.patternForeground(
              pattern.pattern(), toCellColorReport(pattern.foregroundColor()).orElseThrow());
      case ExcelCellFillSnapshot.PatternBackground pattern ->
          CellFillReport.patternBackground(
              pattern.pattern(), toCellColorReport(pattern.backgroundColor()).orElseThrow());
      case ExcelCellFillSnapshot.PatternForegroundBackground pattern ->
          CellFillReport.patternColors(
              pattern.pattern(),
              toCellColorReport(pattern.foregroundColor()).orElseThrow(),
              toCellColorReport(pattern.backgroundColor()).orElseThrow());
      case ExcelCellFillSnapshot.Gradient gradient ->
          CellFillReport.gradient(toCellGradientFillReport(gradient.gradient()));
    };
  }

  private static CellGradientFillReport toCellGradientFillReport(
      ExcelGradientFillSnapshot gradient) {
    return switch (gradient) {
      case ExcelGradientFillSnapshot.Linear linear ->
          CellGradientFillReport.linear(
              linear.degree(),
              linear.stops().stream()
                  .map(
                      stop ->
                          new CellGradientStopReport(
                              stop.position(), toCellColorReport(stop.color()).orElseThrow()))
                  .toList());
      case ExcelGradientFillSnapshot.Path path ->
          CellGradientFillReport.path(
              path.left(),
              path.right(),
              path.top(),
              path.bottom(),
              path.stops().stream()
                  .map(
                      stop ->
                          new CellGradientStopReport(
                              stop.position(), toCellColorReport(stop.color()).orElseThrow()))
                  .toList());
    };
  }
}
