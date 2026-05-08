package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing definition for one border side within a style patch. */
public record CellBorderSideInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ExcelBorderStyle> style,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ColorInput> color) {
  /** Creates a protocol border side with the supplied style and no explicit RGB color. */
  public CellBorderSideInput(ExcelBorderStyle style) {
    this(Optional.ofNullable(style), Optional.empty());
  }

  /** Creates a protocol border side with an explicit optional color payload. */
  public CellBorderSideInput(ExcelBorderStyle style, ColorInput color) {
    this(Optional.ofNullable(style), Optional.ofNullable(color));
  }

  public CellBorderSideInput {
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(color, "color must not be null");
    if (style.isEmpty() && color.isEmpty()) {
      throw new IllegalArgumentException("border side must set style and/or color");
    }
    if (style.isPresent() && style.orElseThrow() == ExcelBorderStyle.NONE && color.isPresent()) {
      throw new IllegalArgumentException("border side color is not supported when style is NONE");
    }
  }
}
