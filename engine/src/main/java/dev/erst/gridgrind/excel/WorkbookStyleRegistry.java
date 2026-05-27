package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;

/**
 * Caches and creates POI CellStyle and Font instances for a single workbook, merging protocol style
 * patches onto existing cell styles.
 */
@SuppressWarnings("PMD.CommentRequired")
public final class WorkbookStyleRegistry {
  private static final ExcelCellStyle LOCAL_DATE_STYLE_PATCH =
      ExcelCellStyle.numberFormat("yyyy-mm-dd");
  private static final ExcelCellStyle LOCAL_DATE_TIME_STYLE_PATCH =
      ExcelCellStyle.numberFormat("yyyy-mm-dd hh:mm:ss");

  private final XSSFWorkbook workbook;
  private final DataFormat dataFormat;
  private final Map<MergedCellStyleKey, XSSFCellStyle> cellStyles;
  private final Map<MergedFontKey, XSSFFont> fonts;
  private final Map<String, Integer> gradientFillIds;
  private final ExcelGradientFillStyleSupport gradientFillSupport;
  private final ExcelBorderPatchSupport borderPatchSupport;
  private final ExcelCellStyleSnapshotSupport snapshotSupport;

  public WorkbookStyleRegistry(XSSFWorkbook workbook) {
    this(workbook, StylesTableFillRegistryAccess.poiApi());
  }

  public WorkbookStyleRegistry(
      XSSFWorkbook workbook, StylesTableFillRegistryAccess fillRegistryAccess) {
    this.workbook = workbook;
    this.dataFormat = workbook.createDataFormat();
    this.cellStyles = new HashMap<>();
    this.fonts = new HashMap<>();
    this.gradientFillIds = new HashMap<>();
    this.gradientFillSupport =
        new ExcelGradientFillStyleSupport(workbook, fillRegistryAccess, gradientFillIds);
    this.borderPatchSupport = new ExcelBorderPatchSupport(workbook);
    this.snapshotSupport = new ExcelCellStyleSnapshotSupport(workbook);
    gradientFillSupport.indexExistingGradientFills();
  }

  /**
   * Returns the current cell style with the local-date number format merged on top.
   *
   * <p>This preserves any existing fill, border, font, alignment, or wrap state already present on
   * the cell.
   */
  public CellStyle localDateStyle(Cell cell) {
    return mergedStyle(cell, LOCAL_DATE_STYLE_PATCH);
  }

  /**
   * Returns the current cell style with the local-date-time number format merged on top.
   *
   * <p>This preserves any existing fill, border, font, alignment, or wrap state already present on
   * the cell.
   */
  public CellStyle localDateTimeStyle(Cell cell) {
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
    return snapshotSupport.snapshot(style);
  }

  /** Captures a read-only snapshot of the workbook's default cell style. */
  ExcelCellStyleSnapshot defaultSnapshot() {
    return snapshotSupport.snapshot(defaultStyleRecord());
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
    ExcelWorkbookStyleLimits.requireCellStyleCapacity(workbook); // LIM-011
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    cellStyle.cloneStyleFrom(baseStyle);

    stylePatch
        .numberFormat()
        .ifPresent(numberFormat -> cellStyle.setDataFormat(dataFormat.getFormat(numberFormat)));
    applyAlignmentPatch(cellStyle, stylePatch.alignment());
    applyFillPatch(cellStyle, stylePatch.fill());
    stylePatch.border().ifPresent(border -> applyBorderPatch(cellStyle, border));
    stylePatch.protection().ifPresent(protection -> applyProtectionPatch(cellStyle, protection));
    stylePatch
        .font()
        .ifPresent(fontPatch -> cellStyle.setFont(fontFor(baseStyle.getFont(), fontPatch)));
    return cellStyle;
  }

  private void applyAlignmentPatch(
      XSSFCellStyle cellStyle, Optional<ExcelCellAlignment> alignmentPatch) {
    if (alignmentPatch.isEmpty()) {
      return;
    }
    ExcelCellAlignment patch = alignmentPatch.orElseThrow();
    patch.wrapText().ifPresent(cellStyle::setWrapText);
    patch
        .horizontalAlignment()
        .ifPresent(value -> cellStyle.setAlignment(ExcelCellStylePoiBridge.toPoi(value)));
    patch
        .verticalAlignment()
        .ifPresent(value -> cellStyle.setVerticalAlignment(ExcelCellStylePoiBridge.toPoi(value)));
    patch.textRotation().ifPresent(value -> cellStyle.setRotation(value.shortValue()));
    patch.indentation().ifPresent(value -> cellStyle.setIndention(value.shortValue()));
  }

  private void applyFillPatch(XSSFCellStyle cellStyle, Optional<ExcelCellFill> fillPatch) {
    if (fillPatch.isEmpty()) {
      return;
    }
    switch (fillPatch.orElseThrow()) {
      case ExcelCellFill.Gradient gradient -> {
        gradientFillSupport.applyGradientFillPatch(cellStyle, gradient.gradient());
        return;
      }
      case ExcelCellFill.PatternOnly pattern -> {
        cellStyle.setFillPattern(ExcelCellStylePoiBridge.toPoi(pattern.pattern()));
        if (pattern.pattern() == ExcelFillPattern.NONE) {
          gradientFillSupport.clearFillColors(cellStyle);
          return;
        }
        if (pattern.pattern() == ExcelFillPattern.SOLID) {
          cellStyle.setFillBackgroundColor((XSSFColor) null);
        }
      }
      case ExcelCellFill.PatternForeground pattern -> {
        cellStyle.setFillPattern(ExcelCellStylePoiBridge.toPoi(pattern.pattern()));
        if (pattern.pattern() == ExcelFillPattern.SOLID) {
          cellStyle.setFillBackgroundColor((XSSFColor) null);
        }
        cellStyle.setFillForegroundColor(
            ExcelColorSupport.toXssfColor(workbook, pattern.foregroundColor()));
      }
      case ExcelCellFill.PatternBackground pattern -> {
        cellStyle.setFillPattern(ExcelCellStylePoiBridge.toPoi(pattern.pattern()));
        cellStyle.setFillBackgroundColor(
            ExcelColorSupport.toXssfColor(workbook, pattern.backgroundColor()));
      }
      case ExcelCellFill.PatternForegroundBackground pattern -> {
        cellStyle.setFillPattern(ExcelCellStylePoiBridge.toPoi(pattern.pattern()));
        cellStyle.setFillForegroundColor(
            ExcelColorSupport.toXssfColor(workbook, pattern.foregroundColor()));
        cellStyle.setFillBackgroundColor(
            ExcelColorSupport.toXssfColor(workbook, pattern.backgroundColor()));
      }
    }
  }

  private void applyProtectionPatch(XSSFCellStyle cellStyle, ExcelCellProtection protectionPatch) {
    protectionPatch.locked().ifPresent(cellStyle::setLocked);
    protectionPatch.hiddenFormula().ifPresent(cellStyle::setHidden);
  }

  private XSSFFont fontFor(XSSFFont baseFont, ExcelCellFont fontPatch) {
    return fonts.computeIfAbsent(
        new MergedFontKey(baseFont.getIndex(), fontPatch),
        key -> createMergedFont(baseFont, fontPatch));
  }

  private XSSFFont createMergedFont(XSSFFont baseFont, ExcelCellFont fontPatch) {
    XSSFFont font = workbook.createFont();
    font.getCTFont().set(baseFont.getCTFont());
    fontPatch.bold().ifPresent(font::setBold);
    fontPatch.italic().ifPresent(font::setItalic);
    fontPatch.fontName().ifPresent(font::setFontName);
    fontPatch
        .fontHeight()
        .ifPresent(fontHeight -> font.setFontHeight(fontHeight.points().doubleValue()));
    fontPatch.fontColor().ifPresent(color -> applyFontColor(font, color));
    fontPatch
        .underline()
        .ifPresent(
            underline -> font.setUnderline(underline ? FontUnderline.SINGLE : FontUnderline.NONE));
    fontPatch.strikeout().ifPresent(font::setStrikeout);
    return font;
  }

  private void applyFontColor(XSSFFont font, ExcelColor color) {
    CTFont ctFont = font.getCTFont();
    while (ctFont.sizeOfColorArray() > 1) {
      ctFont.removeColor(ctFont.sizeOfColorArray() - 1);
    }
    if (ctFont.sizeOfColorArray() == 0) {
      ctFont.addNewColor();
    }
    ctFont.getColorArray(0).set(ExcelColorSupport.toXssfColor(workbook, color).getCTColor());
  }

  private void applyBorderPatch(XSSFCellStyle cellStyle, ExcelBorder border) {
    borderPatchSupport.applyBorderPatch(cellStyle, border);
  }

  ExcelGradientFillSnapshot gradientFillSnapshot(CTGradientFill fill) {
    return snapshotSupport.gradientFillSnapshot(fill);
  }

  private record MergedCellStyleKey(int baseStyleIndex, ExcelCellStyle stylePatch) {}

  private record MergedFontKey(int baseFontIndex, ExcelCellFont fontPatch) {}
}
