package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;

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
    ExcelConditionalFormattingStyleWriteSupport.applyStyle(workbook.getStylesSource(), dxf, style);
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
  static @Nullable ExcelDifferentialStyleSnapshot snapshotStyle(
      StylesTable stylesTable, CTCfRule rule) {
    return ExcelConditionalFormattingStyleSnapshotSupport.optionalSnapshotStyle(stylesTable, rule)
        .orElse(null);
  }

  static FillSnapshot snapshotFill(@Nullable CTFill fill) {
    return ExcelConditionalFormattingStyleSnapshotSupport.snapshotFill(fill);
  }

  /** Reports whether one pattern-fill payload uses a fill pattern GridGrind does not model. */
  static boolean patternTypeIsUnsupported(CTPatternFill patternFill) {
    return ExcelConditionalFormattingStyleSnapshotSupport.patternTypeIsUnsupported(patternFill);
  }

  /**
   * Returns one modeled solid-fill foreground color, recording unsupported fill states as needed.
   */
  static @Nullable ExcelColor patternForegroundColor(
      CTPatternFill patternFill,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    return ExcelConditionalFormattingStyleSnapshotSupport.patternForegroundColor(
        patternFill, unsupportedFeatures);
  }

  /** Reports whether one differential border carries complex features GridGrind does not model. */
  static boolean hasComplexBorderFeatures(CTBorder border) {
    return ExcelConditionalFormattingStyleSnapshotSupport.hasComplexBorderFeatures(border);
  }

  /** Reports whether one differential border references a side payload GridGrind cannot read. */
  static boolean hasUnsupportedSideReference(
      CTBorder border,
      @Nullable ExcelBorderSide top,
      @Nullable ExcelBorderSide right,
      @Nullable ExcelBorderSide bottom,
      @Nullable ExcelBorderSide left) {
    return ExcelConditionalFormattingStyleSnapshotSupport.hasUnsupportedSideReference(
        border, top, right, bottom, left);
  }

  /** Collapses the four explicit sides of one differential border into the public border record. */
  static @Nullable ExcelDifferentialBorder borderValue(
      @Nullable ExcelBorderSide top,
      @Nullable ExcelBorderSide right,
      @Nullable ExcelBorderSide bottom,
      @Nullable ExcelBorderSide left) {
    return ExcelConditionalFormattingStyleSnapshotSupport.borderValue(top, right, bottom, left);
  }

  /** Returns the factual border side modeled by one differential border-side XML payload. */
  static @Nullable ExcelBorderSide snapshotBorderSide(@Nullable CTBorderPr side) {
    return ExcelConditionalFormattingStyleSnapshotSupport.snapshotBorderSide(side);
  }

  /** Returns the effective boolean value represented by one optional OOXML boolean property. */
  static @Nullable Boolean booleanProperty(@Nullable CTBooleanProperty property) {
    return ExcelConditionalFormattingStyleSnapshotSupport.booleanProperty(property);
  }

  /**
   * Returns the effective underline flag represented by one differential-font underline payload.
   */
  static boolean underline(@Nullable CTUnderlineProperty property) {
    return ExcelConditionalFormattingStyleSnapshotSupport.underline(property);
  }

  /** Reports whether one differential-font payload uses font features GridGrind does not model. */
  static boolean hasUnsupportedFontAttributes(CTFont font) {
    return ExcelConditionalFormattingStyleSnapshotSupport.hasUnsupportedFontAttributes(font);
  }

  /** Returns unsupported-feature markers exposed by one raw differential-style metadata payload. */
  static List<ExcelConditionalFormattingUnsupportedFeature> metadataUnsupportedFeatures(CTDxf dxf) {
    return ExcelConditionalFormattingStyleSnapshotSupport.metadataUnsupportedFeatures(dxf);
  }

  /** Converts one GridGrind border-style enum into the matching OOXML border-style constant. */
  static STBorderStyle.Enum toCtBorderStyle(ExcelBorderStyle style) {
    return ExcelConditionalFormattingBorderStyleSupport.toCtBorderStyle(style);
  }

  /** Converts one OOXML border-style code into the matching GridGrind border-style enum. */
  static ExcelBorderStyle fromCtBorderStyle(int styleCode) {
    return ExcelConditionalFormattingBorderStyleSupport.fromCtBorderStyle(styleCode);
  }

  static long putDxf(StylesTable stylesTable, CTDxf dxf) {
    return stylesTable.putDxf(dxf) - 1L;
  }

  static Optional<CTDxf> dxfAt(StylesTable stylesTable, long dxfId) {
    if (dxfId < 0 || dxfId >= stylesTable._getDXfsSize()) {
      return Optional.empty();
    }
    return Optional.of(stylesTable.getDxfAt(Math.toIntExact(dxfId)));
  }

  record FillSnapshot(
      @Nullable ExcelColor fillColor,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    FillSnapshot {
      unsupportedFeatures = List.copyOf(unsupportedFeatures);
    }

    static FillSnapshot empty() {
      return new FillSnapshot(null, List.of());
    }
  }

  record BorderSnapshot(
      @Nullable ExcelDifferentialBorder border,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    BorderSnapshot {
      unsupportedFeatures = List.copyOf(unsupportedFeatures);
    }

    static BorderSnapshot empty() {
      return new BorderSnapshot(null, List.of());
    }
  }

  record FontSnapshot(
      @Nullable Boolean bold,
      @Nullable Boolean italic,
      @Nullable ExcelFontHeight fontHeight,
      @Nullable ExcelColor fontColor,
      @Nullable Boolean underline,
      @Nullable Boolean strikeout,
      List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
    FontSnapshot {
      unsupportedFeatures = List.copyOf(unsupportedFeatures);
    }

    static FontSnapshot empty() {
      return new FontSnapshot(null, null, null, null, null, null, List.of());
    }

    boolean isEmpty() {
      return java.util.stream.Stream.of(bold, italic, fontHeight, fontColor, underline, strikeout)
          .allMatch(java.util.Objects::isNull);
    }
  }
}
