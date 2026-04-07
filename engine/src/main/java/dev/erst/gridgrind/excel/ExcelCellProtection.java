package dev.erst.gridgrind.excel;

/** Protection patch applied through {@link ExcelCellStyle}. */
public record ExcelCellProtection(Boolean locked, Boolean hiddenFormula) {
  public ExcelCellProtection {
    if (locked == null && hiddenFormula == null) {
      throw new IllegalArgumentException("protection must set at least one attribute");
    }
  }
}
