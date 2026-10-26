package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Print-margin payload used by advanced print-layout authoring. */
@CatalogType(
    id = "PrintMarginsInput",
    summary = "Explicit print margins measured in the workbook's stored inch-based values.",
    displayOrder = 64)
public record PrintMarginsInput(
    double left, double right, double top, double bottom, double header, double footer) {
  public PrintMarginsInput {
    requireFinite(left, "left");
    requireFinite(right, "right");
    requireFinite(top, "top");
    requireFinite(bottom, "bottom");
    requireFinite(header, "header");
    requireFinite(footer, "footer");
  }

  private static void requireFinite(double value, String fieldName) {
    if (!Double.isFinite(value) || value < 0.0d) {
      throw new IllegalArgumentException(fieldName + " must be finite and non-negative");
    }
  }
}
