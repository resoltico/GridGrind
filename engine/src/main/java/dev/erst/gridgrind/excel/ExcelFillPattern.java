package dev.erst.gridgrind.excel;

/** Workbook-facing patterned fill choices independent of Apache POI transport types. */
public enum ExcelFillPattern {
  NONE,
  SOLID,
  FINE_DOTS,
  ALT_BARS,
  SPARSE_DOTS,
  THICK_HORIZONTAL_BANDS,
  THICK_VERTICAL_BANDS,
  THICK_BACKWARD_DIAGONAL,
  THICK_FORWARD_DIAGONAL,
  BIG_SPOTS,
  BRICKS,
  THIN_HORIZONTAL_BANDS,
  THIN_VERTICAL_BANDS,
  THIN_BACKWARD_DIAGONAL,
  THIN_FORWARD_DIAGONAL,
  SQUARES,
  DIAMONDS,
  LESS_DOTS,
  LEAST_DOTS
}
