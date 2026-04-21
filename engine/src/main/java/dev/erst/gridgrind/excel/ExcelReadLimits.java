package dev.erst.gridgrind.excel;

/** Canonical engine-owned hard limits for workbook read surfaces. */
public final class ExcelReadLimits {
  /** Maximum number of cells permitted in one rectangular read window. */
  public static final int MAX_WINDOW_CELLS = 250_000; // LIM-001

  private ExcelReadLimits() {}
}
