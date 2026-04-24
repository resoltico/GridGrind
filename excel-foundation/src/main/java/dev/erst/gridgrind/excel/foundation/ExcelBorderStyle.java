package dev.erst.gridgrind.excel.foundation;

/**
 * Border styles that GridGrind can read, analyze, and write for `.xlsx` workbooks.
 *
 * <p>When surfaced through the contract module, these constant names are part of the published wire
 * vocabulary and must be treated as breaking-change-sensitive.
 */
public enum ExcelBorderStyle {
  NONE,
  THIN,
  MEDIUM,
  DASHED,
  DOTTED,
  THICK,
  DOUBLE,
  HAIR,
  MEDIUM_DASHED,
  DASH_DOT,
  MEDIUM_DASH_DOT,
  DASH_DOT_DOT,
  MEDIUM_DASH_DOT_DOT,
  SLANTED_DASH_DOT
}
