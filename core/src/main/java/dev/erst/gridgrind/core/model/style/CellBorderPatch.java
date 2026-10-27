package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing border patch used by {@link CellStylePatch}. */
@CatalogType(
    id = "CellBorderPatch",
    summary =
        "Border patch for cell styling; at least one side must be set. Use 'all' as shorthand for all four sides.",
    displayOrder = 49)
public record CellBorderPatch(
    Optional<BorderSidePatch> all,
    Optional<BorderSidePatch> top,
    Optional<BorderSidePatch> right,
    Optional<BorderSidePatch> bottom,
    Optional<BorderSidePatch> left) {
  public CellBorderPatch {
    Objects.requireNonNull(all, "all must not be null");
    Objects.requireNonNull(top, "top must not be null");
    Objects.requireNonNull(right, "right must not be null");
    Objects.requireNonNull(bottom, "bottom must not be null");
    Objects.requireNonNull(left, "left must not be null");
    if (all.isEmpty() && top.isEmpty() && right.isEmpty() && bottom.isEmpty() && left.isEmpty()) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }
}
