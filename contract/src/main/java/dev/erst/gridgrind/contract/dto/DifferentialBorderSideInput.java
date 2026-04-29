package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing definition for one conditional-formatting differential border side. */
public record DifferentialBorderSideInput(
    ExcelBorderStyle style, @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> color) {
  public DifferentialBorderSideInput {
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(color, "color must not be null");
    color = color.map(value -> ProtocolRgbColorSupport.requireRgbHex(value, "color"));
  }
}
