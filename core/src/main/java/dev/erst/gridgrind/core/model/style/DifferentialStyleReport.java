package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingUnsupportedFeature;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/** Protocol-facing factual report for one conditional-formatting differential style. */
public record DifferentialStyleReport(
    Optional<String> numberFormat,
    Optional<Boolean> bold,
    Optional<Boolean> italic,
    Optional<FontHeightReport> fontHeight,
    Optional<CellColorReport> fontColor,
    Optional<Boolean> underline,
    Optional<Boolean> strikeout,
    Optional<CellColorReport> fillColor,
    Optional<DifferentialBorderReport> border,
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
  public DifferentialStyleReport {
    Objects.requireNonNull(numberFormat, "numberFormat must not be null");
    if (numberFormat.filter(String::isBlank).isPresent()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    Objects.requireNonNull(bold, "bold must not be null");
    Objects.requireNonNull(italic, "italic must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
    Objects.requireNonNull(underline, "underline must not be null");
    Objects.requireNonNull(strikeout, "strikeout must not be null");
    Objects.requireNonNull(fontColor, "fontColor must not be null");
    Objects.requireNonNull(fillColor, "fillColor must not be null");
    Objects.requireNonNull(border, "border must not be null");
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

  private static boolean hasNoStyleAttributes(Optional<?>... attributes) {
    return Stream.of(attributes).allMatch(Optional::isEmpty);
  }
}
