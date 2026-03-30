package dev.erst.gridgrind.excel;

/** Immutable snapshot of a cell after formatting and, when needed, formula evaluation. */
public sealed interface ExcelCellSnapshot {
  /** A1-style address of the cell, such as {@code B4}. */
  String address();

  /** Raw cell type as reported by Excel: STRING, NUMBER, BOOLEAN, FORMULA, BLANK, or ERROR. */
  String declaredType();

  /**
   * Resolved value type after formula evaluation: same as {@link #declaredType()} for non-formula
   * cells; the evaluated result type for formula cells.
   */
  String effectiveType();

  /** Formatted display string as Excel would render it in the cell. */
  String displayValue();

  /** Formatting style applied to the cell at the time of the snapshot. */
  ExcelCellStyleSnapshot style();

  /** Hyperlink and comment metadata captured for the cell at snapshot time. */
  ExcelCellMetadataSnapshot metadata();

  record BlankSnapshot(
      String address,
      String declaredType,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata)
      implements ExcelCellSnapshot {
    @Override
    public String effectiveType() {
      return "BLANK";
    }
  }

  record TextSnapshot(
      String address,
      String declaredType,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      String stringValue)
      implements ExcelCellSnapshot {
    @Override
    public String effectiveType() {
      return "STRING";
    }
  }

  record NumberSnapshot(
      String address,
      String declaredType,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      Double numberValue)
      implements ExcelCellSnapshot {
    @Override
    public String effectiveType() {
      return "NUMBER";
    }
  }

  record BooleanSnapshot(
      String address,
      String declaredType,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      Boolean booleanValue)
      implements ExcelCellSnapshot {
    @Override
    public String effectiveType() {
      return "BOOLEAN";
    }
  }

  record ErrorSnapshot(
      String address,
      String declaredType,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      String errorValue)
      implements ExcelCellSnapshot {
    @Override
    public String effectiveType() {
      return "ERROR";
    }
  }

  record FormulaSnapshot(
      String address,
      String declaredType,
      String displayValue,
      ExcelCellStyleSnapshot style,
      ExcelCellMetadataSnapshot metadata,
      String formula,
      ExcelCellSnapshot evaluation)
      implements ExcelCellSnapshot {
    @Override
    public String effectiveType() {
      return "FORMULA";
    }
  }
}
