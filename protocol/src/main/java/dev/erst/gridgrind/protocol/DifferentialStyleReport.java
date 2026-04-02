package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelConditionalFormattingUnsupportedFeature;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import java.util.List;
import java.util.Objects;
import java.util.stream.Stream;

/** Protocol-facing factual report for one conditional-formatting differential style. */
public record DifferentialStyleReport(
    String numberFormat,
    Boolean bold,
    Boolean italic,
    FontHeightReport fontHeight,
    String fontColor,
    Boolean underline,
    Boolean strikeout,
    String fillColor,
    DifferentialBorderReport border,
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
  public DifferentialStyleReport {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    fontColor = ProtocolRgbColorSupport.normalizeRgbHex(fontColor, "fontColor");
    fillColor = ProtocolRgbColorSupport.normalizeRgbHex(fillColor, "fillColor");
    Objects.requireNonNull(unsupportedFeatures, "unsupportedFeatures must not be null");
    unsupportedFeatures = List.copyOf(unsupportedFeatures);
    for (ExcelConditionalFormattingUnsupportedFeature unsupportedFeature : unsupportedFeatures) {
      Objects.requireNonNull(
          unsupportedFeature, "unsupportedFeatures must not contain null values");
    }
    if (unsupportedFeatures.isEmpty()
        && hasNoStyleAttributes(
            numberFormat,
            bold,
            italic,
            fontHeight,
            fontColor,
            underline,
            strikeout,
            fillColor,
            border)) {
      throw new IllegalArgumentException("style must expose at least one attribute");
    }
  }

  /** Converts one engine differential-style snapshot into the protocol report shape. */
  public static DifferentialStyleReport fromExcel(ExcelDifferentialStyleSnapshot style) {
    if (style == null) {
      return null;
    }
    return new DifferentialStyleReport(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        toFontHeightReport(style),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        DifferentialBorderReport.fromExcel(style.border()),
        style.unsupportedFeatures());
  }

  private static boolean hasNoStyleAttributes(Object... attributes) {
    return Stream.of(attributes).allMatch(Objects::isNull);
  }

  private static FontHeightReport toFontHeightReport(ExcelDifferentialStyleSnapshot style) {
    return style.fontHeight() == null
        ? null
        : FontHeightReport.fromExcelFontHeight(style.fontHeight());
  }
}
