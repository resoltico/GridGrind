package dev.erst.gridgrind.excel;

/** Immutable snapshot of a cell after formatting and, when needed, formula evaluation. */
public sealed interface ExcelCellSnapshot {
  String address();
  String declaredType();
  String effectiveType();
  String displayValue();
  ExcelCellStyleSnapshot style();

  record BlankSnapshot(
      String address, String declaredType, String displayValue, ExcelCellStyleSnapshot style)
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
      Double numberValue)
      implements ExcelCellSnapshot {
    @Override
    public String effectiveType() {
      return "NUMERIC";
    }
  }

  record BooleanSnapshot(
      String address,
      String declaredType,
      String displayValue,
      ExcelCellStyleSnapshot style,
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
      String formula,
      ExcelCellSnapshot evaluation)
      implements ExcelCellSnapshot {
    @Override
    public String effectiveType() {
      return "FORMULA";
    }
  }
}
