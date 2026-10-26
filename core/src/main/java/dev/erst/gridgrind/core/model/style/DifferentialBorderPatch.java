package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/** Protocol-facing differential border patch used by conditional-formatting rule styles. */
@CatalogType(
    id = "DifferentialBorderPatch",
    summary =
        "Conditional-formatting differential border patch; at least one side must be set. Use 'all' as shorthand for all four sides.",
    displayOrder = 61)
public record DifferentialBorderPatch(
    Optional<BorderSidePatch> all,
    Optional<BorderSidePatch> top,
    Optional<BorderSidePatch> right,
    Optional<BorderSidePatch> bottom,
    Optional<BorderSidePatch> left) {
  public DifferentialBorderPatch {
    Objects.requireNonNull(all, "all must not be null");
    Objects.requireNonNull(top, "top must not be null");
    Objects.requireNonNull(right, "right must not be null");
    Objects.requireNonNull(bottom, "bottom must not be null");
    Objects.requireNonNull(left, "left must not be null");
    if (hasNoSides(all, top, right, bottom, left)) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }

  private static boolean hasNoSides(Optional<?>... sides) {
    return Stream.of(sides).allMatch(Optional::isEmpty);
  }
}
