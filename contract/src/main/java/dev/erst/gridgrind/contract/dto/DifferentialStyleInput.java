package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/** Protocol-facing differential style payload for authored conditional-formatting rules. */
public record DifferentialStyleInput(
    String numberFormat,
    Boolean bold,
    Boolean italic,
    FontHeightInput fontHeight,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> fontColor,
    Boolean underline,
    Boolean strikeout,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> fillColor,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialBorderInput> border) {
  public DifferentialStyleInput {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    Objects.requireNonNull(fontColor, "fontColor must not be null");
    Objects.requireNonNull(fillColor, "fillColor must not be null");
    Objects.requireNonNull(border, "border must not be null");
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

  private static boolean hasNoStyleAttributes(Object... attributes) {
    return Stream.of(attributes)
        .map(
            attribute ->
                attribute instanceof Optional<?> optional ? optional.orElse(null) : attribute)
        .allMatch(Objects::isNull);
  }
}
