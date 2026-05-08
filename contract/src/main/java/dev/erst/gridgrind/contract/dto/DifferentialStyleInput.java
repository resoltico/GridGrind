package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing differential style payload for authored conditional-formatting rules. */
public record DifferentialStyleInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> numberFormat,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> bold,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> italic,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<FontHeightInput> fontHeight,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> fontColor,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> underline,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> strikeout,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> fillColor,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialBorderInput> border) {
  public DifferentialStyleInput {
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
    fontColor = fontColor.map(value -> ProtocolRgbColorSupport.requireRgbHex(value, "fontColor"));
    fillColor = fillColor.map(value -> ProtocolRgbColorSupport.requireRgbHex(value, "fillColor"));
    if (hasNoStyleAttributes(
        numberFormat,
        bold,
        italic,
        fontHeight,
        fontColor,
        underline,
        strikeout,
        fillColor,
        border)) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }

  @SafeVarargs
  private static boolean hasNoStyleAttributes(Optional<?>... attributes) {
    return java.util.Arrays.stream(attributes).allMatch(Optional::isEmpty);
  }
}
