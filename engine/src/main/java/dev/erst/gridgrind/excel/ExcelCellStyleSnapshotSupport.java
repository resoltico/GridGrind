package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelGradientFillGeometry;
import java.util.Optional;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientStop;

/** Extracts immutable GridGrind style snapshots from live POI workbook style state. */
final class ExcelCellStyleSnapshotSupport {
  private static final String DEFAULT_NUMBER_FORMAT = "General";

  private final XSSFWorkbook workbook;

  ExcelCellStyleSnapshotSupport(XSSFWorkbook workbook) {
    this.workbook = workbook;
  }

  static ExcelCellFontSnapshot snapshotFont(XSSFFont font) {
    return new ExcelCellFontSnapshot(
        font.getBold(),
        font.getItalic(),
        font.getFontName(),
        new ExcelFontHeight(font.getFontHeight()),
        ExcelColorSnapshotSupport.snapshot(font.getXSSFColor()).orElse(null),
        font.getUnderline() != FontUnderline.NONE.getByteValue(),
        font.getStrikeout());
  }

  static String resolveNumberFormat(String numberFormat) {
    return numberFormat == null || numberFormat.isBlank() ? DEFAULT_NUMBER_FORMAT : numberFormat;
  }

  ExcelCellStyleSnapshot snapshot(XSSFCellStyle style) {
    return new ExcelCellStyleSnapshot(
        resolveNumberFormat(style.getDataFormatString()),
        new ExcelCellAlignmentSnapshot(
            style.getWrapText(),
            ExcelCellStylePoiBridge.fromPoi(style.getAlignment()),
            ExcelCellStylePoiBridge.fromPoi(style.getVerticalAlignment()),
            style.getRotation(),
            style.getIndention()),
        snapshotFont(style.getFont()),
        fillSnapshot(style),
        new ExcelBorderSnapshot(
            borderSideSnapshot(style.getBorderTop(), style.getTopBorderXSSFColor()),
            borderSideSnapshot(style.getBorderRight(), style.getRightBorderXSSFColor()),
            borderSideSnapshot(style.getBorderBottom(), style.getBottomBorderXSSFColor()),
            borderSideSnapshot(style.getBorderLeft(), style.getLeftBorderXSSFColor())),
        new ExcelCellProtectionSnapshot(style.getLocked(), style.getHidden()));
  }

  ExcelGradientFillSnapshot gradientFillSnapshot(CTGradientFill fill) {
    Double left = fill.isSetLeft() ? fill.getLeft() : null;
    Double right = fill.isSetRight() ? fill.getRight() : null;
    Double top = fill.isSetTop() ? fill.getTop() : null;
    Double bottom = fill.isSetBottom() ? fill.getBottom() : null;
    String type =
        ExcelGradientFillGeometry.effectiveType(
            fill.isSetType() ? fill.getType().toString() : null, left, right, top, bottom);
    java.util.List<ExcelGradientStopSnapshot> stops =
        java.util.Arrays.stream(fill.getStopArray()).map(this::gradientStopSnapshot).toList();
    return "PATH".equals(type)
        ? ExcelGradientFillSnapshot.path(left, right, top, bottom, stops)
        : ExcelGradientFillSnapshot.linear(fill.isSetDegree() ? fill.getDegree() : null, stops);
  }

  private ExcelCellFillSnapshot fillSnapshot(XSSFCellStyle style) {
    XSSFCellFill fill = fill(style);
    if (fill.getCTFill().isSetGradientFill()) {
      return ExcelCellFillSnapshot.gradient(
          gradientFillSnapshot(fill.getCTFill().getGradientFill()));
    }
    var pattern = ExcelCellStylePoiBridge.fromPoi(style.getFillPattern());
    if (pattern == dev.erst.gridgrind.excel.foundation.ExcelFillPattern.NONE) {
      return ExcelCellFillSnapshot.pattern(pattern);
    }
    Optional<ExcelColorSnapshot> foreground =
        ExcelColorSnapshotSupport.snapshot(style.getFillForegroundColorColor());
    Optional<ExcelColorSnapshot> background =
        pattern == dev.erst.gridgrind.excel.foundation.ExcelFillPattern.SOLID
            ? Optional.empty()
            : ExcelColorSnapshotSupport.snapshot(style.getFillBackgroundColorColor());
    if (foreground.isPresent() && background.isPresent()) {
      return ExcelCellFillSnapshot.patternColors(
          pattern, foreground.orElseThrow(), background.orElseThrow());
    }
    if (foreground.isPresent()) {
      return ExcelCellFillSnapshot.patternForeground(pattern, foreground.orElseThrow());
    }
    if (background.isPresent()) {
      return ExcelCellFillSnapshot.patternBackground(pattern, background.orElseThrow());
    }
    return ExcelCellFillSnapshot.pattern(pattern);
  }

  private XSSFCellFill fill(XSSFCellStyle style) {
    long fillId = style.getCoreXf().getFillId();
    return workbook.getStylesSource().getFillAt((int) fillId);
  }

  private ExcelGradientStopSnapshot gradientStopSnapshot(CTGradientStop stop) {
    ExcelColorSnapshot color =
        ExcelColorSnapshotSupport.snapshot(workbook, stop.getColor())
            .orElseThrow(
                () ->
                    new IllegalStateException(
                        "Gradient stop at position "
                            + stop.getPosition()
                            + " is missing its color definition."));
    return new ExcelGradientStopSnapshot(stop.getPosition(), color);
  }

  private static ExcelBorderSideSnapshot borderSideSnapshot(
      BorderStyle borderStyle, XSSFColor borderColor) {
    return new ExcelBorderSideSnapshot(
        ExcelCellStylePoiBridge.fromPoi(borderStyle),
        ExcelColorSnapshotSupport.snapshot(borderColor).orElse(null));
  }
}
