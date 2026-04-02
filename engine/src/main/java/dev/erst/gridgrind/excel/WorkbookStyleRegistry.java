package dev.erst.gridgrind.excel;

import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Caches and creates POI CellStyle and Font instances for a single workbook, merging protocol style
 * patches onto existing cell styles.
 */
final class WorkbookStyleRegistry {
  private static final String DEFAULT_NUMBER_FORMAT = "General";
  private static final ExcelCellStyle LOCAL_DATE_STYLE_PATCH =
      ExcelCellStyle.numberFormat("yyyy-mm-dd");
  private static final ExcelCellStyle LOCAL_DATE_TIME_STYLE_PATCH =
      ExcelCellStyle.numberFormat("yyyy-mm-dd hh:mm:ss");

  private final XSSFWorkbook workbook;
  private final DataFormat dataFormat;
  private final Map<MergedCellStyleKey, XSSFCellStyle> cellStyles;
  private final Map<MergedFontKey, XSSFFont> fonts;

  WorkbookStyleRegistry(XSSFWorkbook workbook) {
    this.workbook = workbook;
    this.dataFormat = workbook.createDataFormat();
    this.cellStyles = new HashMap<>();
    this.fonts = new HashMap<>();
  }

  /**
   * Returns the current cell style with the local-date number format merged on top.
   *
   * <p>This preserves any existing fill, border, font, alignment, or wrap state already present on
   * the cell.
   */
  CellStyle localDateStyle(Cell cell) {
    return mergedStyle(cell, LOCAL_DATE_STYLE_PATCH);
  }

  /**
   * Returns the current cell style with the local-date-time number format merged on top.
   *
   * <p>This preserves any existing fill, border, font, alignment, or wrap state already present on
   * the cell.
   */
  CellStyle localDateTimeStyle(Cell cell) {
    return mergedStyle(cell, LOCAL_DATE_TIME_STYLE_PATCH);
  }

  /** Returns the workbook's default cell style (index 0). */
  CellStyle defaultStyle() {
    return defaultStyleRecord();
  }

  /**
   * Resolves the current cell style and merges the provided style patch on top of it, returning a
   * cached or newly created {@link CellStyle}.
   */
  CellStyle mergedStyle(Cell cell, ExcelCellStyle stylePatch) {
    return styleFor(styleRecord(cell), stylePatch);
  }

  /** Captures a read-only snapshot of the effective style applied to the given cell. */
  ExcelCellStyleSnapshot snapshot(Cell cell) {
    XSSFCellStyle style = styleRecord(cell);
    return snapshot(style);
  }

  /** Captures a read-only snapshot of the workbook's default cell style. */
  ExcelCellStyleSnapshot defaultSnapshot() {
    return snapshot(defaultStyleRecord());
  }

  private ExcelCellStyleSnapshot snapshot(XSSFCellStyle style) {
    XSSFFont font = style.getFont();
    return new ExcelCellStyleSnapshot(
        resolveNumberFormat(style.getDataFormatString()),
        font.getBold(),
        font.getItalic(),
        style.getWrapText(),
        fromPoi(style.getAlignment()),
        fromPoi(style.getVerticalAlignment()),
        font.getFontName(),
        new ExcelFontHeight(font.getFontHeight()),
        ExcelRgbColorSupport.toRgbHex(font.getXSSFColor()),
        font.getUnderline() != FontUnderline.NONE.getByteValue(),
        font.getStrikeout(),
        fillColor(style),
        fromPoi(style.getBorderTop()),
        fromPoi(style.getBorderRight()),
        fromPoi(style.getBorderBottom()),
        fromPoi(style.getBorderLeft()));
  }

  /**
   * Returns the number format string, substituting the default "General" format when the raw value
   * is null or blank.
   */
  static String resolveNumberFormat(String numberFormat) {
    return numberFormat == null || numberFormat.isBlank() ? DEFAULT_NUMBER_FORMAT : numberFormat;
  }

  private XSSFCellStyle defaultStyleRecord() {
    return workbook.getCellStyleAt(0);
  }

  private XSSFCellStyle styleRecord(Cell cell) {
    return (XSSFCellStyle) cell.getCellStyle();
  }

  private XSSFCellStyle styleFor(XSSFCellStyle baseStyle, ExcelCellStyle stylePatch) {
    return cellStyles.computeIfAbsent(
        new MergedCellStyleKey(baseStyle.getIndex(), stylePatch),
        key -> createMergedStyle(baseStyle, stylePatch));
  }

  private XSSFCellStyle createMergedStyle(XSSFCellStyle baseStyle, ExcelCellStyle stylePatch) {
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    cellStyle.cloneStyleFrom(baseStyle);

    if (stylePatch.numberFormat() != null) {
      cellStyle.setDataFormat(dataFormat.getFormat(stylePatch.numberFormat()));
    }
    if (stylePatch.wrapText() != null) {
      cellStyle.setWrapText(stylePatch.wrapText());
    }
    if (stylePatch.horizontalAlignment() != null) {
      cellStyle.setAlignment(toPoi(stylePatch.horizontalAlignment()));
    }
    if (stylePatch.verticalAlignment() != null) {
      cellStyle.setVerticalAlignment(toPoi(stylePatch.verticalAlignment()));
    }
    if (stylePatch.fillColor() != null) {
      cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      cellStyle.setFillForegroundColor(
          ExcelRgbColorSupport.toXssfColor(workbook, stylePatch.fillColor()));
    }
    if (stylePatch.border() != null) {
      applyBorderPatch(cellStyle, stylePatch.border());
    }
    if (hasFontChanges(stylePatch)) {
      cellStyle.setFont(fontFor(baseStyle.getFont(), fontPatch(stylePatch)));
    }
    return cellStyle;
  }

  private XSSFFont fontFor(XSSFFont baseFont, FontPatch fontPatch) {
    return fonts.computeIfAbsent(
        new MergedFontKey(baseFont.getIndex(), fontPatch),
        key -> createMergedFont(baseFont, fontPatch));
  }

  private XSSFFont createMergedFont(XSSFFont baseFont, FontPatch fontPatch) {
    XSSFFont font = workbook.createFont();
    font.getCTFont().set(baseFont.getCTFont());
    if (fontPatch.bold() != null) {
      font.setBold(fontPatch.bold());
    }
    if (fontPatch.italic() != null) {
      font.setItalic(fontPatch.italic());
    }
    if (fontPatch.fontName() != null) {
      font.setFontName(fontPatch.fontName());
    }
    if (fontPatch.fontHeight() != null) {
      font.setFontHeight(fontPatch.fontHeight().points().doubleValue());
    }
    if (fontPatch.fontColor() != null) {
      font.setColor(ExcelRgbColorSupport.toXssfColor(workbook, fontPatch.fontColor()));
    }
    if (fontPatch.underline() != null) {
      font.setUnderline(fontPatch.underline() ? FontUnderline.SINGLE : FontUnderline.NONE);
    }
    if (fontPatch.strikeout() != null) {
      font.setStrikeout(fontPatch.strikeout());
    }
    return font;
  }

  private void applyBorderPatch(XSSFCellStyle cellStyle, ExcelBorder border) {
    ExcelBorderStyle topStyle = borderStyle(border.all(), border.top());
    ExcelBorderStyle rightStyle = borderStyle(border.all(), border.right());
    ExcelBorderStyle bottomStyle = borderStyle(border.all(), border.bottom());
    ExcelBorderStyle leftStyle = borderStyle(border.all(), border.left());

    if (topStyle != null) {
      cellStyle.setBorderTop(toPoi(topStyle));
    }
    if (rightStyle != null) {
      cellStyle.setBorderRight(toPoi(rightStyle));
    }
    if (bottomStyle != null) {
      cellStyle.setBorderBottom(toPoi(bottomStyle));
    }
    if (leftStyle != null) {
      cellStyle.setBorderLeft(toPoi(leftStyle));
    }
  }

  private static ExcelBorderStyle borderStyle(
      ExcelBorderSide defaultSide, ExcelBorderSide explicitSide) {
    if (explicitSide != null) {
      return explicitSide.style();
    }
    return defaultSide == null ? null : defaultSide.style();
  }

  private static boolean hasFontChanges(ExcelCellStyle stylePatch) {
    return stylePatch.bold() != null
        || stylePatch.italic() != null
        || stylePatch.fontName() != null
        || stylePatch.fontHeight() != null
        || stylePatch.fontColor() != null
        || stylePatch.underline() != null
        || stylePatch.strikeout() != null;
  }

  private static FontPatch fontPatch(ExcelCellStyle stylePatch) {
    return new FontPatch(
        stylePatch.bold(),
        stylePatch.italic(),
        stylePatch.fontName(),
        stylePatch.fontHeight(),
        stylePatch.fontColor(),
        stylePatch.underline(),
        stylePatch.strikeout());
  }

  private static String fillColor(XSSFCellStyle style) {
    if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
      return null;
    }
    return ExcelRgbColorSupport.toRgbHex(style.getFillForegroundColorColor());
  }

  private static ExcelHorizontalAlignment fromPoi(HorizontalAlignment alignment) {
    return ExcelHorizontalAlignment.valueOf(alignment.name());
  }

  private static ExcelVerticalAlignment fromPoi(VerticalAlignment alignment) {
    return ExcelVerticalAlignment.valueOf(alignment.name());
  }

  private static ExcelBorderStyle fromPoi(BorderStyle borderStyle) {
    return ExcelBorderStyle.valueOf(borderStyle.name());
  }

  private static HorizontalAlignment toPoi(ExcelHorizontalAlignment alignment) {
    return HorizontalAlignment.valueOf(alignment.name());
  }

  private static VerticalAlignment toPoi(ExcelVerticalAlignment alignment) {
    return VerticalAlignment.valueOf(alignment.name());
  }

  private static BorderStyle toPoi(ExcelBorderStyle borderStyle) {
    return BorderStyle.valueOf(borderStyle.name());
  }

  private record MergedCellStyleKey(int baseStyleIndex, ExcelCellStyle stylePatch) {}

  private record MergedFontKey(int baseFontIndex, FontPatch fontPatch) {}

  private record FontPatch(
      Boolean bold,
      Boolean italic,
      String fontName,
      ExcelFontHeight fontHeight,
      String fontColor,
      Boolean underline,
      Boolean strikeout) {}
}
