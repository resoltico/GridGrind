package dev.erst.gridgrind.excel.foundation;

/**
 * Workbook-facing vertical alignment choices independent of Apache POI transport types.
 *
 * <p>When surfaced through the contract module, these constant names are part of the published wire
 * vocabulary and must be treated as breaking-change-sensitive.
 */
public enum ExcelVerticalAlignment {
  TOP,
  CENTER,
  BOTTOM,
  JUSTIFY,
  DISTRIBUTED
}
