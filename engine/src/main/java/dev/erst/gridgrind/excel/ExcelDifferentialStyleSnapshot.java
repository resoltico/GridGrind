package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingUnsupportedFeature;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Factual differential-style metadata loaded from a conditional-formatting rule. */
public record ExcelDifferentialStyleSnapshot(
    @Nullable String numberFormat,
    @Nullable Boolean bold,
    @Nullable Boolean italic,
    @Nullable FontHeightReport fontHeight,
    @Nullable ColorSpec fontColor,
    @Nullable Boolean underline,
    @Nullable Boolean strikeout,
    @Nullable ColorSpec fillColor,
    @Nullable ExcelDifferentialBorder border,
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
  public ExcelDifferentialStyleSnapshot {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    Objects.requireNonNull(unsupportedFeatures, "unsupportedFeatures must not be null");
    unsupportedFeatures = List.copyOf(unsupportedFeatures);
    for (ExcelConditionalFormattingUnsupportedFeature unsupportedFeature : unsupportedFeatures) {
      Objects.requireNonNull(
          unsupportedFeature, "unsupportedFeatures must not contain null values");
    }
    if (java.util.stream.Stream.of(
                numberFormat,
                bold,
                italic,
                fontHeight,
                fontColor,
                underline,
                strikeout,
                fillColor,
                border)
            .allMatch(Objects::isNull)
        && unsupportedFeatures.isEmpty()) {
      throw new IllegalArgumentException("style must expose at least one attribute");
    }
  }
}
