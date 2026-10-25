package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing cell-style patch used for range and cell presentation changes. */
public record CellStylePatchInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> numberFormat,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellAlignmentInput> alignment,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellFontInput> font,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellFillInput> fill,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellBorderInput> border,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CellProtectionInput> protection) {
  public CellStylePatchInput {
    Objects.requireNonNull(numberFormat, "numberFormat must not be null");
    Objects.requireNonNull(alignment, "alignment must not be null");
    Objects.requireNonNull(font, "font must not be null");
    Objects.requireNonNull(fill, "fill must not be null");
    Objects.requireNonNull(border, "border must not be null");
    Objects.requireNonNull(protection, "protection must not be null");
    numberFormat.ifPresent(
        value -> {
          if (value.isBlank()) {
            throw new IllegalArgumentException("numberFormat must not be blank");
          }
        });
    if (numberFormat.isEmpty()
        && alignment.isEmpty()
        && font.isEmpty()
        && fill.isEmpty()
        && border.isEmpty()
        && protection.isEmpty()) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }
}
