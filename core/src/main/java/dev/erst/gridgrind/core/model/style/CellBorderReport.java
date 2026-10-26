package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;

/** Effective border facts reported with every analyzed cell. */
@CatalogType(
    id = "CellBorderReport",
    summary = "Exact four-sided cell-border report.",
    displayOrder = 81)
public record CellBorderReport(
    CellBorderSideReport top,
    CellBorderSideReport right,
    CellBorderSideReport bottom,
    CellBorderSideReport left) {
  public CellBorderReport {
    Objects.requireNonNull(top, "top must not be null");
    Objects.requireNonNull(right, "right must not be null");
    Objects.requireNonNull(bottom, "bottom must not be null");
    Objects.requireNonNull(left, "left must not be null");
  }
}
