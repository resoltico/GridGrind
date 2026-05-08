package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing font patch used by {@link CellStyleInput}. */
public record CellFontInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> bold,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> italic,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> fontName,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<FontHeightInput> fontHeight,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ColorInput> fontColor,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> underline,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Boolean> strikeout) {
  public CellFontInput {
    Objects.requireNonNull(bold, "bold must not be null");
    Objects.requireNonNull(italic, "italic must not be null");
    Objects.requireNonNull(fontName, "fontName must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
    Objects.requireNonNull(fontColor, "fontColor must not be null");
    Objects.requireNonNull(underline, "underline must not be null");
    Objects.requireNonNull(strikeout, "strikeout must not be null");
    fontName.ifPresent(
        value -> {
          if (value.isBlank()) {
            throw new IllegalArgumentException("fontName must not be blank");
          }
        });
    if (bold.isEmpty()
        && italic.isEmpty()
        && fontName.isEmpty()
        && fontHeight.isEmpty()
        && fontColor.isEmpty()
        && underline.isEmpty()
        && strikeout.isEmpty()) {
      throw new IllegalArgumentException("font must set at least one attribute");
    }
  }
}
