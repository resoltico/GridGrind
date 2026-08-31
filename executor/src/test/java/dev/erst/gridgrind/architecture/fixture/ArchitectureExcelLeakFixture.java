package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.excel.ExcelWorkbook;

/** Deliberately invalid public signature that exposes an unexported workbook implementation. */
public final class ArchitectureExcelLeakFixture {
  private ArchitectureExcelLeakFixture() {}

  /** Exposes a workbook implementation solely for architecture-rule regression. */
  public static ExcelWorkbook leakedWorkbook(ExcelWorkbook workbook) {
    return workbook;
  }
}
