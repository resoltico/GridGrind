package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing cell-style patch used for range and cell presentation changes. */
@CatalogType(
    id = "CellStylePatch",
    summary =
        "Style patch applied to a cell or range; at least one field must be set. Colors preserve RGB, theme, indexed, and tint semantics; style subgroups are nested explicitly.",
    displayOrder = 43)
public record CellStylePatch(
    Optional<String> numberFormat,
    Optional<CellAlignmentPatch> alignment,
    Optional<CellFontPatch> font,
    Optional<CellFillSpec> fill,
    Optional<CellBorderPatch> border,
    Optional<CellProtectionPatch> protection) {
  public CellStylePatch {
    Objects.requireNonNull(numberFormat, "numberFormat must not be null");
    Objects.requireNonNull(alignment, "alignment must not be null");
    Objects.requireNonNull(font, "font must not be null");
    Objects.requireNonNull(fill, "fill must not be null");
    Objects.requireNonNull(border, "border must not be null");
    Objects.requireNonNull(protection, "protection must not be null");
    numberFormat.ifPresent(
        value -> {
          TextValidation.requireNonBlank(value, "numberFormat");
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
