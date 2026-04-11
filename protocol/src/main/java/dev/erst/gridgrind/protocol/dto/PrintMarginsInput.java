package dev.erst.gridgrind.protocol.dto;

/** Print-margin payload used by advanced print-layout authoring. */
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
