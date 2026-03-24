package dev.erst.gridgrind.excel;

import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Caches and creates POI CellStyle and Font instances for a single workbook, merging protocol style
 * patches onto existing cell styles.
 */
final class WorkbookStyleRegistry {
  private static final String DEFAULT_NUMBER_FORMAT = "General";

  private final Workbook workbook;
  private final DataFormat dataFormat;
  private final Map<ResolvedCellStyle, CellStyle> cellStyles;
  private final Map<ResolvedFontStyle, Font> fonts;
  private CellStyle localDateStyle;
  private CellStyle localDateTimeStyle;

  WorkbookStyleRegistry(Workbook workbook) {
    this.workbook = workbook;
    this.dataFormat = workbook.createDataFormat();
    this.cellStyles = new HashMap<>();
    this.fonts = new HashMap<>();
  }

  /** Returns the shared date cell style for {@code yyyy-mm-dd} formatted date values. */
  CellStyle localDateStyle() {
    if (localDateStyle == null) {
      localDateStyle =
          styleFor(
              new ResolvedCellStyle(
                  "yyyy-mm-dd",
                  false,
                  false,
                  false,
                  ExcelHorizontalAlignment.GENERAL,
                  ExcelVerticalAlignment.BOTTOM));
    }
    return localDateStyle;
  }

  /** Returns the shared date-time cell style for {@code yyyy-mm-dd hh:mm:ss} formatted values. */
  CellStyle localDateTimeStyle() {
    if (localDateTimeStyle == null) {
      localDateTimeStyle =
          styleFor(
              new ResolvedCellStyle(
                  "yyyy-mm-dd hh:mm:ss",
                  false,
                  false,
                  false,
                  ExcelHorizontalAlignment.GENERAL,
                  ExcelVerticalAlignment.BOTTOM));
    }
    return localDateTimeStyle;
  }

  /** Returns the workbook's default cell style (index 0). */
  CellStyle defaultStyle() {
    return workbook.getCellStyleAt(0);
  }

  /**
   * Resolves the current cell style and merges the provided style patch on top of it, returning a
   * cached or newly created {@link CellStyle}.
   */
  CellStyle mergedStyle(Cell cell, ExcelCellStyle stylePatch) {
    return styleFor(resolvedStyle(cell).merge(stylePatch));
  }

  /** Captures a read-only snapshot of the effective style applied to the given cell. */
  ExcelCellStyleSnapshot snapshot(Cell cell) {
    ResolvedCellStyle style = resolvedStyle(cell);
    return new ExcelCellStyleSnapshot(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        style.wrapText(),
        style.horizontalAlignment().name(),
        style.verticalAlignment().name());
  }

  private ResolvedCellStyle resolvedStyle(Cell cell) {
    return resolvedStyle(cell.getCellStyle());
  }

  private ResolvedCellStyle resolvedStyle(CellStyle cellStyle) {
    Font font = workbook.getFontAt(cellStyle.getFontIndex());
    String numberFormat = cellStyle.getDataFormatString();
    return new ResolvedCellStyle(
        numberFormat == null || numberFormat.isBlank() ? DEFAULT_NUMBER_FORMAT : numberFormat,
        font.getBold(),
        font.getItalic(),
        cellStyle.getWrapText(),
        fromPoi(cellStyle.getAlignment()),
        fromPoi(cellStyle.getVerticalAlignment()));
  }

  private CellStyle styleFor(ResolvedCellStyle style) {
    return cellStyles.computeIfAbsent(
        style,
        key -> {
          CellStyle cellStyle = workbook.createCellStyle();
          cellStyle.setDataFormat(dataFormat.getFormat(key.numberFormat()));
          cellStyle.setWrapText(key.wrapText());
          cellStyle.setAlignment(toPoi(key.horizontalAlignment()));
          cellStyle.setVerticalAlignment(toPoi(key.verticalAlignment()));
          cellStyle.setFont(fontFor(new ResolvedFontStyle(key.bold(), key.italic())));
          return cellStyle;
        });
  }

  private Font fontFor(ResolvedFontStyle fontStyle) {
    return fonts.computeIfAbsent(
        fontStyle,
        key -> {
          Font font = workbook.createFont();
          font.setBold(key.bold());
          font.setItalic(key.italic());
          return font;
        });
  }

  private record ResolvedCellStyle(
      String numberFormat,
      boolean bold,
      boolean italic,
      boolean wrapText,
      ExcelHorizontalAlignment horizontalAlignment,
      ExcelVerticalAlignment verticalAlignment) {
    private ResolvedCellStyle merge(ExcelCellStyle stylePatch) {
      return new ResolvedCellStyle(
          stylePatch.numberFormat() == null ? numberFormat : stylePatch.numberFormat(),
          stylePatch.bold() == null ? bold : stylePatch.bold(),
          stylePatch.italic() == null ? italic : stylePatch.italic(),
          stylePatch.wrapText() == null ? wrapText : stylePatch.wrapText(),
          stylePatch.horizontalAlignment() == null
              ? horizontalAlignment
              : stylePatch.horizontalAlignment(),
          stylePatch.verticalAlignment() == null
              ? verticalAlignment
              : stylePatch.verticalAlignment());
    }
  }

  private record ResolvedFontStyle(boolean bold, boolean italic) {}

  private static ExcelHorizontalAlignment fromPoi(HorizontalAlignment alignment) {
    return ExcelHorizontalAlignment.valueOf(alignment.name());
  }

  private static ExcelVerticalAlignment fromPoi(VerticalAlignment alignment) {
    return ExcelVerticalAlignment.valueOf(alignment.name());
  }

  private static HorizontalAlignment toPoi(ExcelHorizontalAlignment alignment) {
    return HorizontalAlignment.valueOf(alignment.name());
  }

  private static VerticalAlignment toPoi(ExcelVerticalAlignment alignment) {
    return VerticalAlignment.valueOf(alignment.name());
  }
}
