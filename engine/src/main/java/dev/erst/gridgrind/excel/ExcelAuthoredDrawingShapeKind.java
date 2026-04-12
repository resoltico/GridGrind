package dev.erst.gridgrind.excel;

/** Drawing-shape kinds that GridGrind can author during Phase 5. */
public enum ExcelAuthoredDrawingShapeKind {
  /** Author one preset simple shape backed by spreadsheet-drawing geometry. */
  SIMPLE_SHAPE,

  /** Author one connector line anchored by a two-cell drawing box. */
  CONNECTOR
}
