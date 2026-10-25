package dev.erst.gridgrind.contract.dto;

/** Factual worksheet print margins loaded from workbook metadata. */
public record PrintMarginsReport(
    double left, double right, double top, double bottom, double header, double footer) {
  /** Returns the factual worksheet margin defaults for an unconfigured sheet. */
  public static PrintMarginsReport defaults() {
    return new PrintMarginsReport(0.7d, 0.7d, 0.75d, 0.75d, 0.3d, 0.3d);
  }

  public PrintMarginsReport {
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
