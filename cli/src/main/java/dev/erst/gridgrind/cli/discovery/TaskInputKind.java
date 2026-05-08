package dev.erst.gridgrind.cli.discovery;

/** Typed prerequisite inputs required by one CLI-owned task descriptor. */
public enum TaskInputKind {
  SOURCE_WORKBOOK_PATH,
  PERSISTENCE_TARGET_PATH,
  TARGET_SHEET_NAMES,
  TARGET_OBJECT_NAMES,
  CELL_OR_RANGE_COORDINATES,
  TABULAR_SOURCE_ROWS,
  VALIDATION_RULES,
  MAPPING_LOCATOR,
  XML_PAYLOAD,
  BINARY_PAYLOAD,
  DRAWING_ANCHORS
}
