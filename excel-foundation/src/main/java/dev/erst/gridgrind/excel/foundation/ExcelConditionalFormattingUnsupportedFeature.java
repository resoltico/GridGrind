package dev.erst.gridgrind.excel.foundation;

/**
 * Differential-style features detected in workbook content but not modeled directly by GridGrind.
 */
public enum ExcelConditionalFormattingUnsupportedFeature {
  STYLE_REFERENCE,
  FONT_ATTRIBUTES,
  FILL_PATTERN,
  FILL_BACKGROUND_COLOR,
  BORDER_COMPLEXITY,
  ALIGNMENT,
  PROTECTION
}
