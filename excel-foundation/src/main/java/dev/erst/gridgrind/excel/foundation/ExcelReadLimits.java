package dev.erst.gridgrind.excel.foundation;

/** Canonical engine-owned hard limits for workbook read surfaces. */
public final class ExcelReadLimits {
  /** Maximum number of factual cells permitted in one cell-returning read surface. */
  public static final int MAX_READ_CELLS = 250_000; // LIM-001

  private ExcelReadLimits() {}
}
