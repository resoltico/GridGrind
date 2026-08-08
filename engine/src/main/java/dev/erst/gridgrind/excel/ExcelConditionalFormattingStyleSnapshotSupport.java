package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.ExcelConditionalFormattingStyleSupport.BorderSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingStyleSupport.FillSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingStyleSupport.FontSnapshot;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.model.StylesTable;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues;

/** Reads OOXML differential-style payloads back into GridGrind snapshot shapes. */
final class ExcelConditionalFormattingStyleSnapshotSupport {
  private ExcelConditionalFormattingStyleSnapshotSupport() {}

  static Optional<ExcelDifferentialStyleSnapshot> optionalSnapshotStyle(
      StylesTable stylesTable, CTCfRule rule) {
    Objects.requireNonNull(stylesTable, "stylesTable must not be null");
    Objects.requireNonNull(rule, "rule must not be null");
    if (!rule.isSetDxfId()) {
      return Optional.empty();
    }

    CTDxf dxf =
        ExcelConditionalFormattingStyleSupport.dxfAt(stylesTable, rule.getDxfId()).orElse(null);
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

  static FillSnapshot snapshotFill(@Nullable CTFill fill) {
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

  static boolean patternTypeIsUnsupported(CTPatternFill patternFill) {
    return patternFill.isSetPatternType()
        && patternFill.getPatternType() != STPatternType.SOLID
        && patternFill.getPatternType() != STPatternType.NONE;
  }

  static @Nullable ExcelColor patternForegroundColor(
      CTPatternFill patternFill,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    ExcelColor fillColor = optionalPatternForegroundColor(patternFill).orElse(null);
    if (fillColor == null
        && (patternTypeIsUnsupported(patternFill) || patternFill.isSetFgColor())) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FILL_PATTERN);
    }
    return fillColor;
  }

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

  static boolean hasUnsupportedSideReference(
      CTBorder border,
      @Nullable ExcelBorderSide top,
      @Nullable ExcelBorderSide right,
      @Nullable ExcelBorderSide bottom,
      @Nullable ExcelBorderSide left) {
    return java.util.stream.Stream.of(
            border.isSetTop() && top == null,
            border.isSetRight() && right == null,
            border.isSetBottom() && bottom == null,
            border.isSetLeft() && left == null)
        .anyMatch(Boolean.TRUE::equals);
  }

  static @Nullable ExcelDifferentialBorder borderValue(
      @Nullable ExcelBorderSide top,
      @Nullable ExcelBorderSide right,
      @Nullable ExcelBorderSide bottom,
      @Nullable ExcelBorderSide left) {
    return optionalBorderValue(top, right, bottom, left).orElse(null);
  }

  static @Nullable ExcelBorderSide snapshotBorderSide(@Nullable CTBorderPr side) {
    return optionalSnapshotBorderSide(side).orElse(null);
  }

  static @Nullable Boolean booleanProperty(@Nullable CTBooleanProperty property) {
    return property == null ? null : !property.isSetVal() || property.getVal();
  }

  static boolean underline(@Nullable CTUnderlineProperty property) {
    if (property == null || !property.isSetVal()) {
      return true;
    }
    return property.getVal() != STUnderlineValues.NONE;
  }

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

  private static Optional<ExcelDifferentialStyleSnapshot> optionalStyleSnapshot(
      @Nullable String numberFormat,
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

  private static FontSnapshot snapshotFont(@Nullable CTFont font) {
    if (font == null) {
      return FontSnapshot.empty();
    }

    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    ExcelColor fontColor = fontColor(font, unsupportedFeatures);
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

  private static @Nullable ExcelColor fontColor(
      CTFont font, List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    ExcelColor fontColor = optionalFontColor(font).orElse(null);
    if (fontColor == null && font.sizeOfColorArray() > 0) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES);
    }
    return fontColor;
  }

  private static boolean hasGradientFill(CTFill fill) {
    return fill.xmlText().contains("gradientFill");
  }

  private static BorderSnapshot snapshotBorder(@Nullable CTBorder border) {
    if (border == null) {
      return BorderSnapshot.empty();
    }

    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures = new ArrayList<>();
    if (hasComplexBorderFeatures(border)) {
      unsupportedFeatures.add(ExcelConditionalFormattingUnsupportedFeature.BORDER_COMPLEXITY);
    }

    ExcelBorderSide top = borderSide(border.isSetTop(), border::getTop);
    ExcelBorderSide right = borderSide(border.isSetRight(), border::getRight);
    ExcelBorderSide bottom = borderSide(border.isSetBottom(), border::getBottom);
    ExcelBorderSide left = borderSide(border.isSetLeft(), border::getLeft);

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

  private static @Nullable ExcelBorderSide borderSide(
      boolean present, java.util.function.Supplier<CTBorderPr> sideSupplier) {
    return present ? snapshotBorderSide(sideSupplier.get()) : null;
  }

  private static Optional<ExcelColor> optionalFontColor(CTFont font) {
    return font.sizeOfColorArray() == 0
        ? Optional.empty()
        : ExcelConditionalFormattingColorSupport.optionalColor(font.getColorArray(0));
  }

  private static Optional<ExcelColor> optionalPatternForegroundColor(CTPatternFill patternFill) {
    if (patternTypeIsUnsupported(patternFill) || !patternFill.isSetFgColor()) {
      return Optional.empty();
    }
    return ExcelConditionalFormattingColorSupport.optionalColor(patternFill.getFgColor());
  }

  private static Optional<ExcelDifferentialBorder> optionalBorderValue(
      @Nullable ExcelBorderSide top,
      @Nullable ExcelBorderSide right,
      @Nullable ExcelBorderSide bottom,
      @Nullable ExcelBorderSide left) {
    if (top == null && right == null && bottom == null && left == null) {
      return Optional.empty();
    }
    return Optional.of(new ExcelDifferentialBorder(null, top, right, bottom, left));
  }

  private static Optional<ExcelBorderSide> optionalSnapshotBorderSide(@Nullable CTBorderPr side) {
    if (side == null) {
      return Optional.empty();
    }
    if (!side.isSetStyle() && !side.isSetColor()) {
      return Optional.empty();
    }
    Optional<ExcelBorderStyle> style =
        side.isSetStyle()
            ? Optional.of(
                ExcelConditionalFormattingBorderStyleSupport.fromCtBorderStyle(
                    side.getStyle().intValue()))
            : Optional.empty();
    Optional<ExcelColor> color =
        side.isSetColor()
            ? ExcelConditionalFormattingColorSupport.optionalColor(side.getColor())
            : Optional.empty();
    if (side.isSetColor() && color.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(new ExcelBorderSide(style, color));
  }

  @SafeVarargs
  private static List<ExcelConditionalFormattingUnsupportedFeature> normalizedUnsupportedFeatures(
      List<ExcelConditionalFormattingUnsupportedFeature>... groups) {
    java.util.Set<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        new LinkedHashSet<>();
    for (List<ExcelConditionalFormattingUnsupportedFeature> group : groups) {
      unsupportedFeatures.addAll(group);
    }
    return List.copyOf(unsupportedFeatures);
  }
}
