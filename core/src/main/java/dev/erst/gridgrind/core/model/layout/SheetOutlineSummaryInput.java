package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Outline-summary placement authored as part of sheet-presentation state. */
@CatalogType(
    id = "SheetOutlineSummaryInput",
    summary =
        "Outline-summary placement authored as part of sheet-presentation state. Supply both placement flags explicitly.",
    displayOrder = 68)
public record SheetOutlineSummaryInput(boolean rowSumsBelow, boolean rowSumsRight) {
  /** Returns the effective Excel defaults for outline summary placement. */
  public static SheetOutlineSummaryInput defaults() {
    return new SheetOutlineSummaryInput(true, true);
  }

  public SheetOutlineSummaryInput {}
}
