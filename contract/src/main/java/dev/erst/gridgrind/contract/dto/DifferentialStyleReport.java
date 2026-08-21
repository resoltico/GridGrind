package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;
import org.jspecify.annotations.Nullable;

/** Protocol-facing factual report for one conditional-formatting differential style. */
public record DifferentialStyleReport(
    @Nullable String numberFormat,
    @Nullable Boolean bold,
    @Nullable Boolean italic,
    @Nullable FontHeightReport fontHeight,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellColorReport> fontColor,
    @Nullable Boolean underline,
    @Nullable Boolean strikeout,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellColorReport> fillColor,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialBorderReport> border,
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures) {
  public DifferentialStyleReport {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
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

  private static boolean hasNoStyleAttributes(@Nullable Object... attributes) {
    return Stream.of(attributes)
        .map(
            attribute ->
                attribute instanceof Optional<?> optional ? optional.orElse(null) : attribute)
        .allMatch(Objects::isNull);
  }
}
