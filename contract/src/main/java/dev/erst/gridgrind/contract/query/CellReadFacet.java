package dev.erst.gridgrind.contract.query;

/** Optional factual facets that one cell-returning read may project. */
public enum CellReadFacet {
  VALUE,
  STYLE,
  FORMAT,
  HYPERLINK,
  COMMENT,
  FORMULA,
  RICH_TEXT_RUNS,
  TEMPORAL
}
