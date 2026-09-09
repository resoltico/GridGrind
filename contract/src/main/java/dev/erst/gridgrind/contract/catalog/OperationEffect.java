package dev.erst.gridgrind.contract.catalog;

/** One static workbook or external-resource effect an operation may require. */
public enum OperationEffect {
  READ_WORKBOOK,
  MUTATE_WORKBOOK,
  CELLS,
  FORMULAS,
  STYLES,
  HYPERLINKS,
  COMMENTS,
  DRAWINGS,
  WORKBOOK_STRUCTURE,
  SHEET_STRUCTURE,
  ROWS,
  COLUMNS,
  TABLES,
  NAMED_RANGES,
  DATA_VALIDATIONS,
  CONDITIONAL_FORMATTING,
  AUTOFILTERS,
  PIVOT_TABLES,
  CUSTOM_XML,
  WORKBOOK_METADATA,
  WORKBOOK_SECURITY
}
