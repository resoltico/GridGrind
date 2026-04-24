package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFontSize;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTNumFmt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues;

/**
 * Reads and writes conditional-formatting differential styles through the workbook styles table.
 */
final class ExcelConditionalFormattingStyleSupport {
  private ExcelConditionalFormattingStyleSupport() {}

  /** Writes one authored differential style onto the supplied conditional-formatting rule XML. */
  static void applyStyle(XSSFWorkbook workbook, CTCfRule rule, ExcelDifferentialStyle style) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(style, "style must not be null");

    CTDxf dxf = CTDxf.Factory.newInstance();
    applyNumberFormat(workbook.getStylesSource(), dxf, style.numberFormat());
    applyFont(dxf, style);
    applyFill(dxf, style.fillColor());
    applyBorder(dxf, style.border());
    attachStyle(workbook.getStylesSource(), rule, dxf);
  }

  /**
   * Attaches one raw differential-style XML payload to the supplied conditional-formatting rule.
   */
  static void attachStyle(StylesTable stylesTable, CTCfRule rule, CTDxf dxf) {
    Objects.requireNonNull(stylesTable, "stylesTable must not be null");
    Objects.requireNonNull(rule, "rule must not be null");
    Objects.requireNonNull(dxf, "dxf must not be null");
    rule.setDxfId(putDxf(stylesTable, dxf));
  }

  /**
   * Returns the factual differential-style snapshot attached to one conditional-formatting rule.
   */
  static ExcelDifferentialStyleSnapshot snapshotStyle(StylesTable stylesTable, CTCfRule rule) {
    return optionalSnapshotStyle(stylesTable, rule).orElse(null);
  }

  private static Optional<ExcelDifferentialStyleSnapshot> optionalSnapshotStyle(
      StylesTable stylesTable, CTCfRule rule) {
    Objects.requireNonNull(stylesTable, "stylesTable must not be null");
    Objects.requireNonNull(rule, "rule must not be null");
    if (!rule.isSetDxfId()) {
      return Optional.empty();
    }

    CTDxf dxf = dxfAt(stylesTable, rule.getDxfId()).orElse(null);
    if (dxf == null) {
      return Optional.of(
          new ExcelDifferentialStyleSnapshot(
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              null,
              List.of(ExcelConditionalFormattingUnsupportedFeature.STYLE_REFERENCE)));
    }

    FontSnapshot font = snapshotFont(dxf.isSetFont() ? dxf.getFont() : null);
    FillSnapshot fill = snapshotFill(dxf.isSetFill() ? dxf.getFill() : null);
    BorderSnapshot border = snapshotBorder(dxf.isSetBorder() ? dxf.getBorder() : null);
    return optionalStyleSnapshot(
        dxf.isSetNumFmt() ? dxf.getNumFmt().getFormatCode() : null,
        font,
        fill,
        border,
        metadataUnsupportedFeatures(dxf));
  }

  private static long putDxf(StylesTable stylesTable, CTDxf dxf) {
    return stylesTable.putDxf(dxf) - 1L;
  }

  private static Optional<CTDxf> dxfAt(StylesTable stylesTable, long dxfId) {
    if (dxfId < 0 || dxfId >= stylesTable._getDXfsSize()) {
      return Optional.empty();
    }
    return Optional.of(stylesTable.getDxfAt(Math.toIntExact(dxfId)));
  }

  private static void applyNumberFormat(StylesTable stylesTable, CTDxf dxf, String numberFormat) {
    if (numberFormat == null) {
      return;
    }
    int formatId = stylesTable.putNumberFormat(numberFormat);
    CTNumFmt numFmt = dxf.addNewNumFmt();
    numFmt.setNumFmtId(formatId);
    numFmt.setFormatCode(numberFormat);
  }

  private static void applyFont(CTDxf dxf, ExcelDifferentialStyle style) {
    if (java.util.stream.Stream.of(
            style.bold(),
            style.italic(),
            style.fontHeight(),
            style.fontColor(),
            style.underline(),
            style.strikeout())
        .allMatch(Objects::isNull)) {
      return;
    }

    CTFont font = dxf.addNewFont();
    setBooleanProperty(style.bold(), font::addNewB);
    setBooleanProperty(style.italic(), font::addNewI);
    setBooleanProperty(style.strikeout(), font::addNewStrike);
    if (style.fontHeight() != null) {
      CTFontSize fontSize = font.addNewSz();
      fontSize.setVal(style.fontHeight().points().doubleValue());
    }
    if (style.fontColor() != null) {
      setColor(font.addNewColor(), style.fontColor());
    }
    if (style.underline() != null) {
      CTUnderlineProperty underlineProperty = font.addNewU();
      underlineProperty.setVal(
          style.underline() ? STUnderlineValues.SINGLE : STUnderlineValues.NONE);
    }
  }

  private static void applyFill(CTDxf dxf, String fillColor) {
    if (fillColor == null) {
      return;
    }
    CTFill fill = dxf.addNewFill();
    CTPatternFill patternFill = fill.addNewPatternFill();
    patternFill.setPatternType(STPatternType.SOLID);
    setColor(patternFill.addNewFgColor(), fillColor);
  }

  private static void applyBorder(CTDxf dxf, ExcelDifferentialBorder border) {
    if (border == null) {
      return;
    }
    CTBorder ctBorder = dxf.addNewBorder();
    applyBorderSide(ctBorder::addNewTop, resolvedSide(border.all(), border.top()));
    applyBorderSide(ctBorder::addNewRight, resolvedSide(border.all(), border.right()));
    applyBorderSide(ctBorder::addNewBottom, resolvedSide(border.all(), border.bottom()));
    applyBorderSide(ctBorder::addNewLeft, resolvedSide(border.all(), border.left()));
  }

  private static void applyBorderSide(
      java.util.function.Supplier<CTBorderPr> borderPrFactory, ExcelDifferentialBorderSide side) {
    Objects.requireNonNull(borderPrFactory, "borderPrFactory must not be null");
    if (side == null) {
      return;
    }
    CTBorderPr borderPr = borderPrFactory.get();
    borderPr.setStyle(toCtBorderStyle(side.style()));
    if (side.color() != null) {
      setColor(borderPr.addNewColor(), side.color());
    }
  }

  private static ExcelDifferentialBorderSide resolvedSide(
      ExcelDifferentialBorderSide defaultSide, ExcelDifferentialBorderSide explicitSide) {
    return explicitSide == null ? defaultSide : explicitSide;
  }

  private static Optional<ExcelDifferentialStyleSnapshot> optionalStyleSnapshot(
      String numberFormat,
      FontSnapshot font,
      FillSnapshot fill,
      BorderSnapshot border,
      List<ExcelConditionalFormattingUnsupportedFeature> metadataUnsupportedFeatures) {
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        normalizedUnsupportedFeatures(
            font.unsupportedFeatures(),
            fill.unsupportedFeatures(),
            border.unsupportedFeatures(),
            metadataUnsupportedFeatures);
    if (java.util.stream.Stream.of(numberFormat, fill.fillColor(), border.border())
            .allMatch(Objects::isNull)
        && font.isEmpty()
        && unsupportedFeatures.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelDifferentialStyleSnapshot(
            numberFormat,
            font.bold(),
            font.italic(),
            font.fontHeight(),
            font.fontColor(),
            font.underline(),
            font.strikeout(),
            fill.fillColor(),
            border.border(),
            unsupportedFeatures));
  }

  private static FontSnapshot snapshotFont(CTFont font) {
    if (font == null) {
      return FontSnapshot.empty();
    }

    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    String fontColor = fontColor(font, unsupportedFeatures);
    if (hasUnsupportedFontAttributes(font)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES);
    }

    return new FontSnapshot(
        booleanProperty(font.sizeOfBArray() > 0 ? font.getBArray(0) : null),
        booleanProperty(font.sizeOfIArray() > 0 ? font.getIArray(0) : null),
        fontHeight(font).orElse(null),
        fontColor,
        font.sizeOfUArray() > 0 ? underline(font.getUArray(0)) : null,
        booleanProperty(font.sizeOfStrikeArray() > 0 ? font.getStrikeArray(0) : null),
        normalizedUnsupportedFeatures(unsupportedFeatures));
  }

  private static Optional<ExcelFontHeight> fontHeight(CTFont font) {
    if (font.sizeOfSzArray() == 0) {
      return Optional.empty();
    }
    return Optional.of(ExcelFontHeight.fromPoints(BigDecimal.valueOf(font.getSzArray(0).getVal())));
  }

  private static String fontColor(
      CTFont font, List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    String fontColor = optionalFontColor(font).orElse(null);
    if (fontColor == null && font.sizeOfColorArray() > 0) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES);
    }
    return fontColor;
  }

  /** Returns the factual fill metadata modeled by one differential fill payload. */
  static FillSnapshot snapshotFill(CTFill fill) {
    if (fill == null) {
      return FillSnapshot.empty();
    }

    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (hasGradientFill(fill)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN);
    }
    if (!fill.isSetPatternFill()) {
      return new FillSnapshot(null, normalizedUnsupportedFeatures(unsupportedFeatures));
    }

    CTPatternFill patternFill = fill.getPatternFill();
    unsupportedFeatures.addAll(patternFillUnsupportedFeatures(patternFill));
    return new FillSnapshot(
        patternForegroundColor(patternFill, unsupportedFeatures),
        normalizedUnsupportedFeatures(unsupportedFeatures));
  }

  private static boolean hasGradientFill(CTFill fill) {
    return fill.xmlText().contains("gradientFill");
  }

  private static BorderSnapshot snapshotBorder(CTBorder border) {
    if (border == null) {
      return BorderSnapshot.empty();
    }

    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (hasComplexBorderFeatures(border)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY);
    }

    ExcelDifferentialBorderSide top = borderSide(border.isSetTop(), border::getTop);
    ExcelDifferentialBorderSide right = borderSide(border.isSetRight(), border::getRight);
    ExcelDifferentialBorderSide bottom = borderSide(border.isSetBottom(), border::getBottom);
    ExcelDifferentialBorderSide left = borderSide(border.isSetLeft(), border::getLeft);

    if (hasUnsupportedSideReference(border, top, right, bottom, left)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY);
    }

    return new BorderSnapshot(
        borderValue(top, right, bottom, left), normalizedUnsupportedFeatures(unsupportedFeatures));
  }

  private static List<ExcelConditionalFormattingUnsupportedFeature> patternFillUnsupportedFeatures(
      CTPatternFill patternFill) {
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (patternFill.isSetBgColor()) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_BACKGROUND_COLOR);
    }
    if (patternTypeIsUnsupported(patternFill)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN);
    }
    return List.copyOf(unsupportedFeatures);
  }

  /** Reports whether one pattern-fill payload uses a fill pattern GridGrind does not model. */
  static boolean patternTypeIsUnsupported(CTPatternFill patternFill) {
    return patternFill.isSetPatternType()
        && patternFill.getPatternType() != STPatternType.SOLID
        && patternFill.getPatternType() != STPatternType.NONE;
  }

  /**
   * Returns one modeled solid-fill foreground color, recording unsupported fill states as needed.
   */
  static String patternForegroundColor(
      CTPatternFill patternFill,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    String fillColor = optionalPatternForegroundColor(patternFill).orElse(null);
    if (fillColor == null
        && (patternTypeIsUnsupported(patternFill) || patternFill.isSetFgColor())) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN);
    }
    return fillColor;
  }

  /** Reports whether one differential border carries complex features GridGrind does not model. */
  static boolean hasComplexBorderFeatures(CTBorder border) {
    return java.util.stream.Stream.of(
            border.isSetDiagonal(),
            border.isSetVertical(),
            border.isSetHorizontal(),
            border.isSetStart(),
            border.isSetEnd(),
            border.isSetDiagonalDown(),
            border.isSetDiagonalUp())
        .anyMatch(Boolean.TRUE::equals);
  }

  /** Reports whether one differential border references a side payload GridGrind cannot read. */
  static boolean hasUnsupportedSideReference(
      CTBorder border,
      ExcelDifferentialBorderSide top,
      ExcelDifferentialBorderSide right,
      ExcelDifferentialBorderSide bottom,
      ExcelDifferentialBorderSide left) {
    return java.util.stream.Stream.of(
            border.isSetTop() && top == null,
            border.isSetRight() && right == null,
            border.isSetBottom() && bottom == null,
            border.isSetLeft() && left == null)
        .anyMatch(Boolean.TRUE::equals);
  }

  private static ExcelDifferentialBorderSide borderSide(
      boolean present, java.util.function.Supplier<CTBorderPr> sideSupplier) {
    return present ? snapshotBorderSide(sideSupplier.get()) : null;
  }

  /** Collapses the four explicit sides of one differential border into the public border record. */
  static ExcelDifferentialBorder borderValue(
      ExcelDifferentialBorderSide top,
      ExcelDifferentialBorderSide right,
      ExcelDifferentialBorderSide bottom,
      ExcelDifferentialBorderSide left) {
    return optionalBorderValue(top, right, bottom, left).orElse(null);
  }

  /** Returns the factual border side modeled by one differential border-side XML payload. */
  static ExcelDifferentialBorderSide snapshotBorderSide(CTBorderPr side) {
    return optionalSnapshotBorderSide(side).orElse(null);
  }

  /** Returns the effective boolean value represented by one optional OOXML boolean property. */
  static Boolean booleanProperty(CTBooleanProperty property) {
    return property == null ? null : !property.isSetVal() || property.getVal();
  }

  /**
   * Returns the effective underline flag represented by one differential-font underline payload.
   */
  static boolean underline(CTUnderlineProperty property) {
    if (property == null || !property.isSetVal()) {
      return true;
    }
    return property.getVal() != STUnderlineValues.NONE;
  }

  /** Reports whether one differential-font payload uses font features GridGrind does not model. */
  static boolean hasUnsupportedFontAttributes(CTFont font) {
    return java.util.stream.IntStream.of(
            font.sizeOfNameArray(),
            font.sizeOfCharsetArray(),
            font.sizeOfFamilyArray(),
            font.sizeOfOutlineArray(),
            font.sizeOfShadowArray(),
            font.sizeOfCondenseArray(),
            font.sizeOfExtendArray(),
            font.sizeOfVertAlignArray(),
            font.sizeOfSchemeArray())
        .anyMatch(size -> size > 0);
  }

  private static void setBooleanProperty(
      Boolean value, java.util.function.Supplier<CTBooleanProperty> propertyFactory) {
    if (value == null) {
      return;
    }
    CTBooleanProperty property = propertyFactory.get();
    property.setVal(value);
  }

  private static void setColor(CTColor color, String rgbHex) {
    Objects.requireNonNull(color, "color must not be null");
    color.setRgb(argbBytes(rgbHex));
  }

  static byte[] argbBytes(String rgbHex) {
    String normalized =
        ExcelRgbColorSupport.normalizeRgbHex(rgbHex, "rgbHex")
            .orElseThrow(() -> new IllegalArgumentException("rgbHex must not be null"));
    return new byte[] {
      (byte) 0xFF,
      (byte) Integer.parseInt(normalized.substring(1, 3), 16),
      (byte) Integer.parseInt(normalized.substring(3, 5), 16),
      (byte) Integer.parseInt(normalized.substring(5, 7), 16)
    };
  }

  /** Converts one raw OOXML color payload into {@code #RRGGBB}, or null when RGB is unavailable. */
  static String rgbHexFromCtColor(CTColor color) {
    return optionalRgbHexFromCtColor(color).orElse(null);
  }

  private static Optional<String> optionalFontColor(CTFont font) {
    return font.sizeOfColorArray() == 0
        ? Optional.empty()
        : optionalRgbHexFromCtColor(font.getColorArray(0));
  }

  private static Optional<String> optionalPatternForegroundColor(CTPatternFill patternFill) {
    if (patternTypeIsUnsupported(patternFill) || !patternFill.isSetFgColor()) {
      return Optional.empty();
    }
    return optionalRgbHexFromCtColor(patternFill.getFgColor());
  }

  private static Optional<ExcelDifferentialBorder> optionalBorderValue(
      ExcelDifferentialBorderSide top,
      ExcelDifferentialBorderSide right,
      ExcelDifferentialBorderSide bottom,
      ExcelDifferentialBorderSide left) {
    if (top == null && right == null && bottom == null && left == null) {
      return Optional.empty();
    }
    return Optional.of(new ExcelDifferentialBorder(null, top, right, bottom, left));
  }

  private static Optional<ExcelDifferentialBorderSide> optionalSnapshotBorderSide(CTBorderPr side) {
    if (side == null) {
      return Optional.empty();
    }
    if (!side.isSetStyle() && !side.isSetColor()) {
      return Optional.empty();
    }
    ExcelBorderStyle style =
        side.isSetStyle() ? fromCtBorderStyle(side.getStyle().intValue()) : ExcelBorderStyle.NONE;
    String color = side.isSetColor() ? rgbHexFromCtColor(side.getColor()) : null;
    if (side.isSetColor() && color == null) {
      return Optional.empty();
    }
    return Optional.of(new ExcelDifferentialBorderSide(style, color));
  }

  private static Optional<String> optionalRgbHexFromCtColor(CTColor color) {
    if (color == null || !color.isSetRgb()) {
      return Optional.empty();
    }
    byte[] rgb = color.getRgb();
    if (rgb.length == 4) {
      return Optional.of("#%02X%02X%02X".formatted(rgb[1] & 0xFF, rgb[2] & 0xFF, rgb[3] & 0xFF));
    }
    if (rgb.length == 3) {
      return Optional.of("#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
    }
    return Optional.empty();
  }

  /** Returns unsupported-feature markers exposed by one raw differential-style metadata payload. */
  static List<ExcelConditionalFormattingUnsupportedFeature> metadataUnsupportedFeatures(CTDxf dxf) {
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (dxf.isSetAlignment()) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT);
    }
    if (dxf.isSetProtection()) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.PROTECTION);
    }
    return List.copyOf(unsupportedFeatures);
  }

  @SafeVarargs
  private static List<ExcelConditionalFormattingUnsupportedFeature> normalizedUnsupportedFeatures(
      List<ExcelConditionalFormattingUnsupportedFeature>... groups) {
    java.util.Set<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        new java.util.LinkedHashSet<>();
    for (List<ExcelConditionalFormattingUnsupportedFeature> group : groups) {
      unsupportedFeatures.addAll(group);
    }
    return List.copyOf(unsupportedFeatures);
  }

  /** Converts one GridGrind border-style enum into the matching OOXML border-style constant. */
  static STBorderStyle.Enum toCtBorderStyle(ExcelBorderStyle style) {
    return switch (style) {
      case NONE -> STBorderStyle.NONE;
      case THIN -> STBorderStyle.THIN;
      case MEDIUM -> STBorderStyle.MEDIUM;
      case DASHED -> STBorderStyle.DASHED;
      case DOTTED -> STBorderStyle.DOTTED;
      case THICK -> STBorderStyle.THICK;
      case DOUBLE -> STBorderStyle.DOUBLE;
      case HAIR -> STBorderStyle.HAIR;
      case MEDIUM_DASHED -> STBorderStyle.MEDIUM_DASHED;
      case DASH_DOT -> STBorderStyle.DASH_DOT;
      case MEDIUM_DASH_DOT -> STBorderStyle.MEDIUM_DASH_DOT;
      case DASH_DOT_DOT -> STBorderStyle.DASH_DOT_DOT;
      case MEDIUM_DASH_DOT_DOT -> STBorderStyle.MEDIUM_DASH_DOT_DOT;
      case SLANTED_DASH_DOT -> STBorderStyle.SLANT_DASH_DOT;
    };
  }

  /** Converts one OOXML border-style code into the matching GridGrind border-style enum. */
  static ExcelBorderStyle fromCtBorderStyle(int styleCode) {
    return switch (styleCode) {
      case STBorderStyle.INT_NONE -> ExcelBorderStyle.NONE;
      case STBorderStyle.INT_THIN -> ExcelBorderStyle.THIN;
      case STBorderStyle.INT_MEDIUM -> ExcelBorderStyle.MEDIUM;
      case STBorderStyle.INT_DASHED -> ExcelBorderStyle.DASHED;
      case STBorderStyle.INT_DOTTED -> ExcelBorderStyle.DOTTED;
      case STBorderStyle.INT_THICK -> ExcelBorderStyle.THICK;
      case STBorderStyle.INT_DOUBLE -> ExcelBorderStyle.DOUBLE;
      case STBorderStyle.INT_HAIR -> ExcelBorderStyle.HAIR;
      case STBorderStyle.INT_MEDIUM_DASHED -> ExcelBorderStyle.MEDIUM_DASHED;
      case STBorderStyle.INT_DASH_DOT -> ExcelBorderStyle.DASH_DOT;
      case STBorderStyle.INT_MEDIUM_DASH_DOT -> ExcelBorderStyle.MEDIUM_DASH_DOT;
      case STBorderStyle.INT_DASH_DOT_DOT -> ExcelBorderStyle.DASH_DOT_DOT;
      case STBorderStyle.INT_MEDIUM_DASH_DOT_DOT -> ExcelBorderStyle.MEDIUM_DASH_DOT_DOT;
      case STBorderStyle.INT_SLANT_DASH_DOT -> ExcelBorderStyle.SLANTED_DASH_DOT;
      default -> throw new IllegalArgumentException("Unsupported CT border style: " + styleCode);
    };
  }

  record FillSnapshot(
      String fillColor, List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    private static FillSnapshot empty() {
      return new FillSnapshot(null, List.of());
    }
  }

  private record BorderSnapshot(
      ExcelDifferentialBorder border,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    private static BorderSnapshot empty() {
      return new BorderSnapshot(null, List.of());
    }
  }

  private record FontSnapshot(
      Boolean bold,
      Boolean italic,
      ExcelFontHeight fontHeight,
      String fontColor,
      Boolean underline,
      Boolean strikeout,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    private static FontSnapshot empty() {
      return new FontSnapshot(null, null, null, null, null, null, List.of());
    }

    private boolean isEmpty() {
      return java.util.stream.Stream.of(bold, italic, fontHeight, fontColor, underline, strikeout)
          .allMatch(java.util.Objects::isNull);
    }
  }
}
