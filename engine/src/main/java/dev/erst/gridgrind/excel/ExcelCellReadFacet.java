package dev.erst.gridgrind.excel;

/** Internal read facets carried by workbook read commands and results. */
public enum ExcelCellReadFacet {
  VALUE,
  STYLE,
  FORMAT,
  HYPERLINK,
  COMMENT,
  FORMULA,
  RICH_TEXT_RUNS,
  TEMPORAL
}
