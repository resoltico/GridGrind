package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing differential style payload for authored conditional-formatting rules. */
@CatalogType(
    id = "DifferentialStylePatch",
    summary =
        "Differential style payload used by authored conditional-formatting rules. At least one field must be set. Colors preserve RGB, theme, indexed, and tint semantics.",
    displayOrder = 60)
public record DifferentialStylePatch(
    Optional<String> numberFormat,
    Optional<Boolean> bold,
    Optional<Boolean> italic,
    Optional<FontHeightSpec> fontHeight,
    Optional<ColorSpec> fontColor,
    Optional<Boolean> underline,
    Optional<Boolean> strikeout,
    Optional<ColorSpec> fillColor,
    Optional<DifferentialBorderPatch> border) {
  public DifferentialStylePatch {
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
