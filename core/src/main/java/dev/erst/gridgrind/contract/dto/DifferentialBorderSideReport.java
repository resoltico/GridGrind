package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing factual report for one conditional-formatting differential border side. */
public record DifferentialBorderSideReport(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<ExcelBorderStyle> style,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellColorReport> color) {
  public DifferentialBorderSideReport {
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
