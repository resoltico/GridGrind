package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelGradientFillGeometry;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
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
  private final Map<String, Integer> gradientFillIds;
  private final StylesTableFillRegistryAccess fillRegistryAccess;

  WorkbookStyleRegistry(XSSFWorkbook workbook) {
    this(workbook, StylesTableFillRegistryAccess.poiApi());
  }

  WorkbookStyleRegistry(XSSFWorkbook workbook, StylesTableFillRegistryAccess fillRegistryAccess) {
    this.workbook = workbook;
    this.dataFormat = workbook.createDataFormat();
    this.cellStyles = new HashMap<>();
    this.fonts = new HashMap<>();
    this.gradientFillIds = new HashMap<>();
    this.fillRegistryAccess = fillRegistryAccess;
    indexExistingGradientFills();
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
        ExcelColorSnapshotSupport.snapshot(font.getXSSFColor()).orElse(null),
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
    patch.horizontalAlignment().ifPresent(value -> cellStyle.setAlignment(toPoi(value)));
    patch.verticalAlignment().ifPresent(value -> cellStyle.setVerticalAlignment(toPoi(value)));
    patch.textRotation().ifPresent(value -> cellStyle.setRotation(value.shortValue()));
    patch.indentation().ifPresent(value -> cellStyle.setIndention(value.shortValue()));
  }

  private void applyFillPatch(XSSFCellStyle cellStyle, Optional<ExcelCellFill> fillPatch) {
    if (fillPatch.isEmpty()) {
      return;
    }
    switch (fillPatch.orElseThrow()) {
      case ExcelCellFill.Gradient gradient -> {
        applyGradientFillPatch(cellStyle, gradient.gradient());
        return;
      }
      case ExcelCellFill.PatternOnly pattern -> {
        cellStyle.setFillPattern(toPoi(pattern.pattern()));
        if (pattern.pattern() == ExcelFillPattern.NONE) {
          clearFillColors(cellStyle);
          return;
        }
        if (pattern.pattern() == ExcelFillPattern.SOLID) {
          cellStyle.setFillBackgroundColor((XSSFColor) null);
        }
      }
      case ExcelCellFill.PatternForeground pattern -> {
        cellStyle.setFillPattern(toPoi(pattern.pattern()));
        if (pattern.pattern() == ExcelFillPattern.SOLID) {
          cellStyle.setFillBackgroundColor((XSSFColor) null);
        }
        cellStyle.setFillForegroundColor(
            ExcelColorSupport.toXssfColor(workbook, pattern.foregroundColor()));
      }
      case ExcelCellFill.PatternBackground pattern -> {
        cellStyle.setFillPattern(toPoi(pattern.pattern()));
        cellStyle.setFillBackgroundColor(
            ExcelColorSupport.toXssfColor(workbook, pattern.backgroundColor()));
      }
      case ExcelCellFill.PatternForegroundBackground pattern -> {
        cellStyle.setFillPattern(toPoi(pattern.pattern()));
        cellStyle.setFillForegroundColor(
            ExcelColorSupport.toXssfColor(workbook, pattern.foregroundColor()));
        cellStyle.setFillBackgroundColor(
            ExcelColorSupport.toXssfColor(workbook, pattern.backgroundColor()));
      }
    }
  }

  private void applyGradientFillPatch(XSSFCellStyle cellStyle, ExcelGradientFill gradientPatch) {
    CTFill gradientFill = CTFill.Factory.newInstance();
    CTGradientFill gradient = gradientFill.addNewGradientFill();
    switch (gradientPatch) {
      case ExcelGradientFill.Path path -> {
        gradient.setType(
            org.openxmlformats.schemas.spreadsheetml.x2006.main.STGradientType.Enum.forString(
                "path"));
        path.left().ifPresent(gradient::setLeft);
        path.right().ifPresent(gradient::setRight);
        path.top().ifPresent(gradient::setTop);
        path.bottom().ifPresent(gradient::setBottom);
      }
      case ExcelGradientFill.Linear linear -> linear.degree().ifPresent(gradient::setDegree);
    }
    for (ExcelGradientStop stop : gradientPatch.stops()) {
      CTGradientStop ctStop = gradient.addNewStop();
      ctStop.setPosition(stop.position());
      ctStop.addNewColor().set(ExcelColorSupport.toXssfColor(workbook, stop.color()).getCTColor());
    }
    int gradientFillId = gradientFillId(gradientFill);
    cellStyle.getCoreXf().setApplyFill(true);
    cellStyle.getCoreXf().setFillId(gradientFillId);
  }

  private int gradientFillId(CTFill gradientFill) {
    String key = gradientFill.xmlText();
    Integer existingId = gradientFillIds.get(key);
    if (existingId != null) {
      return existingId;
    }
    int fillId =
        appendFill(new XSSFCellFill(gradientFill, workbook.getStylesSource().getIndexedColors()));
    gradientFillIds.put(key, fillId);
    return fillId;
  }

  private void indexExistingGradientFills() {
    List<XSSFCellFill> fills = fillsList();
    for (int fillId = 0; fillId < fills.size(); fillId++) {
      XSSFCellFill fill = fills.get(fillId);
      if (fill.getCTFill().isSetGradientFill()) {
        gradientFillIds.putIfAbsent(fill.getCTFill().xmlText(), fillId);
      }
    }
  }

  private int appendFill(XSSFCellFill fill) {
    return fillRegistryAccess.appendFill(workbook.getStylesSource(), fill);
  }

  private List<XSSFCellFill> fillsList() {
    return fillRegistryAccess.fills(workbook.getStylesSource());
  }

  private void clearFillColors(XSSFCellStyle cellStyle) {
    cellStyle.setFillForegroundColor((XSSFColor) null);
    cellStyle.setFillBackgroundColor((XSSFColor) null);
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

  private Optional<ExcelBorderSide> mergedBorderSide(
      Optional<ExcelBorderSide> defaultSide, Optional<ExcelBorderSide> explicitSide) {
    return effectiveBorderSide(
        mergedBorderStyle(defaultSide, explicitSide), mergedBorderColor(defaultSide, explicitSide));
  }

  private Optional<ExcelBorderStyle> mergedBorderStyle(
      Optional<ExcelBorderSide> defaultSide, Optional<ExcelBorderSide> explicitSide) {
    Objects.requireNonNull(defaultSide, "defaultSide must not be null");
    Objects.requireNonNull(explicitSide, "explicitSide must not be null");
    Optional<ExcelBorderStyle> explicitStyle = explicitSide.flatMap(ExcelBorderSide::style);
    return explicitStyle.isPresent() ? explicitStyle : defaultSide.flatMap(ExcelBorderSide::style);
  }

  private Optional<ExcelColor> mergedBorderColor(
      Optional<ExcelBorderSide> defaultSide, Optional<ExcelBorderSide> explicitSide) {
    Objects.requireNonNull(defaultSide, "defaultSide must not be null");
    Objects.requireNonNull(explicitSide, "explicitSide must not be null");
    if (explicitSide.isPresent() && explicitSide.orElseThrow().color().isPresent()) {
      return explicitSide.orElseThrow().color();
    }
    return defaultSide.isEmpty() || defaultSide.orElseThrow().color().isEmpty()
        ? Optional.empty()
        : defaultSide.orElseThrow().color();
  }

  private Optional<ExcelBorderSide> effectiveBorderSide(
      Optional<ExcelBorderStyle> style, Optional<ExcelColor> color) {
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(color, "color must not be null");
    if (style.isEmpty() && color.isEmpty()) {
      return Optional.empty();
    }
    if (color.isPresent() && (style.isEmpty() || style.orElseThrow() == ExcelBorderStyle.NONE)) {
      throw new IllegalArgumentException("border side color requires an effective border style");
    }
    return Optional.of(new ExcelBorderSide(style, color));
  }

  private void applyBorderSidePatch(
      Optional<ExcelBorderSide> sidePatch,
      Consumer<BorderStyle> styleSetter,
      Consumer<XSSFColor> colorSetter) {
    Objects.requireNonNull(sidePatch, "sidePatch must not be null");
    if (sidePatch.isEmpty()) {
      return;
    }
    ExcelBorderSide resolved = sidePatch.orElseThrow();
    styleSetter.accept(toPoi(resolved.style().orElseThrow()));
    if (resolved.style().orElseThrow() == ExcelBorderStyle.NONE) {
      // POI clears the side color as part of resetting the border style to NONE. An additional
      // explicit null-color clear is redundant and can crash when the XML <color> child is absent.
      return;
    }
    if (resolved.color().isPresent()) {
      colorSetter.accept(ExcelColorSupport.toXssfColor(workbook, resolved.color().orElseThrow()));
    }
  }

  private ExcelCellFillSnapshot fillSnapshot(XSSFCellStyle style) {
    XSSFCellFill fill = fill(style);
    if (fill.getCTFill().isSetGradientFill()) {
      return ExcelCellFillSnapshot.gradient(
          gradientFillSnapshot(fill.getCTFill().getGradientFill()));
    }
    ExcelFillPattern pattern = fromPoi(style.getFillPattern());
    if (pattern == ExcelFillPattern.NONE) {
      return ExcelCellFillSnapshot.pattern(pattern);
    }
    Optional<ExcelColorSnapshot> foreground =
        ExcelColorSnapshotSupport.snapshot(style.getFillForegroundColorColor());
    Optional<ExcelColorSnapshot> background =
        pattern == ExcelFillPattern.SOLID
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
    ExcelBorderStyle style = fromPoi(borderStyle);
    return new ExcelBorderSideSnapshot(
        style, ExcelColorSnapshotSupport.snapshot(borderColor).orElse(null));
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
