package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** Authorable conditional-formatting differential style patch. */
public record ExcelDifferentialStyle(
    Optional<String> numberFormat,
    Optional<Boolean> bold,
    Optional<Boolean> italic,
    Optional<ExcelFontHeight> fontHeight,
    Optional<ExcelColor> fontColor,
    Optional<Boolean> underline,
    Optional<Boolean> strikeout,
    Optional<ExcelColor> fillColor,
    Optional<ExcelDifferentialBorder> border) {
  public ExcelDifferentialStyle {
    Objects.requireNonNull(numberFormat, "numberFormat must not be null");
    Objects.requireNonNull(bold, "bold must not be null");
    Objects.requireNonNull(italic, "italic must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
    Objects.requireNonNull(fontColor, "fontColor must not be null");
    Objects.requireNonNull(underline, "underline must not be null");
    Objects.requireNonNull(strikeout, "strikeout must not be null");
    Objects.requireNonNull(fillColor, "fillColor must not be null");
    Objects.requireNonNull(border, "border must not be null");
    numberFormat =
        numberFormat.map(
            value -> {
              if (value.isBlank()) {
                throw new IllegalArgumentException("numberFormat must not be blank");
              }
              return value;
            });
    if (java.util.stream.Stream.of(
            numberFormat.orElse(null),
            bold.orElse(null),
            italic.orElse(null),
            fontHeight.orElse(null),
            fontColor.orElse(null),
            underline.orElse(null),
            strikeout.orElse(null),
            fillColor.orElse(null),
            border.orElse(null))
        .allMatch(Objects::isNull)) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }
}
