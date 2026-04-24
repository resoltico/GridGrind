package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
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
    fontColor = ProtocolRgbColorSupport.normalizeRgbHex(fontColor, "fontColor").orElse(null);
    fillColor = ProtocolRgbColorSupport.normalizeRgbHex(fillColor, "fillColor").orElse(null);
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

  private static boolean hasNoStyleAttributes(Object... attributes) {
    return Stream.of(attributes).allMatch(Objects::isNull);
  }
}
