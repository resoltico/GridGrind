package dev.erst.gridgrind.excel;

import java.util.HashMap;
import java.util.Map;
import java.util.function.Consumer;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientStop;

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

  /** Captures a factual snapshot of one POI font, including theme-resolved RGB color. */
  static ExcelCellFontSnapshot snapshotFont(XSSFFont font) {
    return new ExcelCellFontSnapshot(
        font.getBold(),
        font.getItalic(),
        font.getFontName(),
        new ExcelFontHeight(font.getFontHeight()),
        ExcelColorSnapshotSupport.snapshot(font.getXSSFColor()),
        font.getUnderline() != FontUnderline.NONE.getByteValue(),
        font.getStrikeout());
  }

  private ExcelCellStyleSnapshot snapshot(XSSFCellStyle style) {
    return new ExcelCellStyleSnapshot(
        resolveNumberFormat(style.getDataFormatString()),
        new ExcelCellAlignmentSnapshot(
            style.getWrapText(),
            fromPoi(style.getAlignment()),
            fromPoi(style.getVerticalAlignment()),
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
    applyAlignmentPatch(cellStyle, stylePatch.alignment());
    applyFillPatch(cellStyle, stylePatch.fill());
    if (stylePatch.border() != null) {
      applyBorderPatch(cellStyle, stylePatch.border());
    }
    if (stylePatch.protection() != null) {
      applyProtectionPatch(cellStyle, stylePatch.protection());
    }
    if (stylePatch.font() != null) {
      cellStyle.setFont(fontFor(baseStyle.getFont(), stylePatch.font()));
    }
    return cellStyle;
  }

  private void applyAlignmentPatch(XSSFCellStyle cellStyle, ExcelCellAlignment alignmentPatch) {
    if (alignmentPatch == null) {
      return;
    }
    if (alignmentPatch.wrapText() != null) {
      cellStyle.setWrapText(alignmentPatch.wrapText());
    }
    if (alignmentPatch.horizontalAlignment() != null) {
      cellStyle.setAlignment(toPoi(alignmentPatch.horizontalAlignment()));
    }
    if (alignmentPatch.verticalAlignment() != null) {
      cellStyle.setVerticalAlignment(toPoi(alignmentPatch.verticalAlignment()));
    }
    if (alignmentPatch.textRotation() != null) {
      cellStyle.setRotation(alignmentPatch.textRotation().shortValue());
    }
    if (alignmentPatch.indentation() != null) {
      cellStyle.setIndention(alignmentPatch.indentation().shortValue());
    }
  }

  private void applyFillPatch(XSSFCellStyle cellStyle, ExcelCellFill fillPatch) {
    if (fillPatch == null) {
      return;
    }

    ExcelFillPattern effectivePattern = effectiveFillPattern(cellStyle, fillPatch);
    if (effectivePattern != null) {
      cellStyle.setFillPattern(toPoi(effectivePattern));
      if (effectivePattern == ExcelFillPattern.NONE) {
        clearFillColors(cellStyle);
        return;
      }
      if (effectivePattern == ExcelFillPattern.SOLID) {
        cellStyle.setFillBackgroundColor((XSSFColor) null);
      }
    }
    if (fillPatch.foregroundColor() != null) {
      cellStyle.setFillForegroundColor(
          ExcelRgbColorSupport.toXssfColor(workbook, fillPatch.foregroundColor()));
    }
    if (fillPatch.backgroundColor() != null) {
      cellStyle.setFillBackgroundColor(
          ExcelRgbColorSupport.toXssfColor(workbook, fillPatch.backgroundColor()));
    }
  }

  private void clearFillColors(XSSFCellStyle cellStyle) {
    cellStyle.setFillForegroundColor((XSSFColor) null);
    cellStyle.setFillBackgroundColor((XSSFColor) null);
  }

  private ExcelFillPattern effectiveFillPattern(XSSFCellStyle cellStyle, ExcelCellFill fillPatch) {
    if (fillPatch.pattern() != null) {
      return fillPatch.pattern();
    }
    if (fromPoi(cellStyle.getFillPattern()) == ExcelFillPattern.NONE) {
      return ExcelFillPattern.SOLID;
    }
    return null;
  }

  private void applyProtectionPatch(XSSFCellStyle cellStyle, ExcelCellProtection protectionPatch) {
    if (protectionPatch.locked() != null) {
      cellStyle.setLocked(protectionPatch.locked());
    }
    if (protectionPatch.hiddenFormula() != null) {
      cellStyle.setHidden(protectionPatch.hiddenFormula());
    }
  }

  private XSSFFont fontFor(XSSFFont baseFont, ExcelCellFont fontPatch) {
    return fonts.computeIfAbsent(
        new MergedFontKey(baseFont.getIndex(), fontPatch),
        key -> createMergedFont(baseFont, fontPatch));
  }

  private XSSFFont createMergedFont(XSSFFont baseFont, ExcelCellFont fontPatch) {
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
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.top()),
        cellStyle::setBorderTop,
        cellStyle::setTopBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.right()),
        cellStyle::setBorderRight,
        cellStyle::setRightBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.bottom()),
        cellStyle::setBorderBottom,
        cellStyle::setBottomBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.left()),
        cellStyle::setBorderLeft,
        cellStyle::setLeftBorderColor);
  }

  private ExcelBorderSide mergedBorderSide(
      ExcelBorderSide defaultSide, ExcelBorderSide explicitSide) {
    return effectiveBorderSide(
        mergedBorderStyle(defaultSide, explicitSide), mergedBorderColor(defaultSide, explicitSide));
  }

  private ExcelBorderStyle mergedBorderStyle(
      ExcelBorderSide defaultSide, ExcelBorderSide explicitSide) {
    if (explicitSide != null && explicitSide.style() != null) {
      return explicitSide.style();
    }
    return defaultSide != null ? defaultSide.style() : null;
  }

  private String mergedBorderColor(ExcelBorderSide defaultSide, ExcelBorderSide explicitSide) {
    if (explicitSide != null && explicitSide.color() != null) {
      return explicitSide.color();
    }
    return defaultSide != null ? defaultSide.color() : null;
  }

  private ExcelBorderSide effectiveBorderSide(ExcelBorderStyle style, String color) {
    if (style == null && color == null) {
      return null;
    }
    if (color != null && (style == null || style == ExcelBorderStyle.NONE)) {
      throw new IllegalArgumentException("border side color requires an effective border style");
    }
    return new ExcelBorderSide(style, color);
  }

  private void applyBorderSidePatch(
      ExcelBorderSide sidePatch,
      Consumer<BorderStyle> styleSetter,
      Consumer<XSSFColor> colorSetter) {
    if (sidePatch == null) {
      return;
    }
    styleSetter.accept(toPoi(sidePatch.style()));
    if (sidePatch.style() == ExcelBorderStyle.NONE) {
      // POI clears the side color as part of resetting the border style to NONE. An additional
      // explicit null-color clear is redundant and can crash when the XML <color> child is absent.
      return;
    }
    if (sidePatch.color() != null) {
      colorSetter.accept(ExcelRgbColorSupport.toXssfColor(workbook, sidePatch.color()));
    }
  }

  private ExcelCellFillSnapshot fillSnapshot(XSSFCellStyle style) {
    XSSFCellFill fill = fill(style);
    if (fill.getCTFill().isSetGradientFill()) {
      return new ExcelCellFillSnapshot(
          ExcelFillPattern.NONE,
          null,
          null,
          gradientFillSnapshot(fill.getCTFill().getGradientFill()));
    }
    ExcelFillPattern pattern = fromPoi(style.getFillPattern());
    if (pattern == ExcelFillPattern.NONE) {
      return new ExcelCellFillSnapshot(pattern, null, null, null);
    }
    return new ExcelCellFillSnapshot(
        pattern,
        ExcelColorSnapshotSupport.snapshot(style.getFillForegroundColorColor()),
        pattern == ExcelFillPattern.SOLID
            ? null
            : ExcelColorSnapshotSupport.snapshot(style.getFillBackgroundColorColor()),
        null);
  }

  private XSSFCellFill fill(XSSFCellStyle style) {
    long fillId = style.getCoreXf().getFillId();
    return workbook.getStylesSource().getFillAt((int) fillId);
  }

  ExcelGradientFillSnapshot gradientFillSnapshot(CTGradientFill fill) {
    String type = fill.isSetType() ? fill.getType().toString() : "LINEAR";
    java.util.List<ExcelGradientStopSnapshot> stops =
        java.util.Arrays.stream(fill.getStopArray()).map(this::gradientStopSnapshot).toList();
    return new ExcelGradientFillSnapshot(
        type,
        fill.isSetDegree() ? fill.getDegree() : null,
        fill.isSetLeft() ? fill.getLeft() : null,
        fill.isSetRight() ? fill.getRight() : null,
        fill.isSetTop() ? fill.getTop() : null,
        fill.isSetBottom() ? fill.getBottom() : null,
        stops);
  }

  private ExcelGradientStopSnapshot gradientStopSnapshot(CTGradientStop stop) {
    return new ExcelGradientStopSnapshot(
        stop.getPosition(), ExcelColorSnapshotSupport.snapshot(workbook, stop.getColor()));
  }

  private static ExcelBorderSideSnapshot borderSideSnapshot(
      BorderStyle borderStyle, XSSFColor borderColor) {
    ExcelBorderStyle style = fromPoi(borderStyle);
    return new ExcelBorderSideSnapshot(style, ExcelColorSnapshotSupport.snapshot(borderColor));
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

  private static ExcelFillPattern fromPoi(FillPatternType pattern) {
    return switch (pattern) {
      case NO_FILL -> ExcelFillPattern.NONE;
      case SOLID_FOREGROUND -> ExcelFillPattern.SOLID;
      case FINE_DOTS -> ExcelFillPattern.FINE_DOTS;
      case ALT_BARS -> ExcelFillPattern.ALT_BARS;
      case SPARSE_DOTS -> ExcelFillPattern.SPARSE_DOTS;
      case THICK_HORZ_BANDS -> ExcelFillPattern.THICK_HORIZONTAL_BANDS;
      case THICK_VERT_BANDS -> ExcelFillPattern.THICK_VERTICAL_BANDS;
      case THICK_BACKWARD_DIAG -> ExcelFillPattern.THICK_BACKWARD_DIAGONAL;
      case THICK_FORWARD_DIAG -> ExcelFillPattern.THICK_FORWARD_DIAGONAL;
      case BIG_SPOTS -> ExcelFillPattern.BIG_SPOTS;
      case BRICKS -> ExcelFillPattern.BRICKS;
      case THIN_HORZ_BANDS -> ExcelFillPattern.THIN_HORIZONTAL_BANDS;
      case THIN_VERT_BANDS -> ExcelFillPattern.THIN_VERTICAL_BANDS;
      case THIN_BACKWARD_DIAG -> ExcelFillPattern.THIN_BACKWARD_DIAGONAL;
      case THIN_FORWARD_DIAG -> ExcelFillPattern.THIN_FORWARD_DIAGONAL;
      case SQUARES -> ExcelFillPattern.SQUARES;
      case DIAMONDS -> ExcelFillPattern.DIAMONDS;
      case LESS_DOTS -> ExcelFillPattern.LESS_DOTS;
      case LEAST_DOTS -> ExcelFillPattern.LEAST_DOTS;
    };
  }

  private static FillPatternType toPoi(ExcelFillPattern pattern) {
    return switch (pattern) {
      case NONE -> FillPatternType.NO_FILL;
      case SOLID -> FillPatternType.SOLID_FOREGROUND;
      case FINE_DOTS -> FillPatternType.FINE_DOTS;
      case ALT_BARS -> FillPatternType.ALT_BARS;
      case SPARSE_DOTS -> FillPatternType.SPARSE_DOTS;
      case THICK_HORIZONTAL_BANDS -> FillPatternType.THICK_HORZ_BANDS;
      case THICK_VERTICAL_BANDS -> FillPatternType.THICK_VERT_BANDS;
      case THICK_BACKWARD_DIAGONAL -> FillPatternType.THICK_BACKWARD_DIAG;
      case THICK_FORWARD_DIAGONAL -> FillPatternType.THICK_FORWARD_DIAG;
      case BIG_SPOTS -> FillPatternType.BIG_SPOTS;
      case BRICKS -> FillPatternType.BRICKS;
      case THIN_HORIZONTAL_BANDS -> FillPatternType.THIN_HORZ_BANDS;
      case THIN_VERTICAL_BANDS -> FillPatternType.THIN_VERT_BANDS;
      case THIN_BACKWARD_DIAGONAL -> FillPatternType.THIN_BACKWARD_DIAG;
      case THIN_FORWARD_DIAGONAL -> FillPatternType.THIN_FORWARD_DIAG;
      case SQUARES -> FillPatternType.SQUARES;
      case DIAMONDS -> FillPatternType.DIAMONDS;
      case LESS_DOTS -> FillPatternType.LESS_DOTS;
      case LEAST_DOTS -> FillPatternType.LEAST_DOTS;
    };
  }

  private record MergedCellStyleKey(int baseStyleIndex, ExcelCellStyle stylePatch) {}

  private record MergedFontKey(int baseFontIndex, ExcelCellFont fontPatch) {}
}
