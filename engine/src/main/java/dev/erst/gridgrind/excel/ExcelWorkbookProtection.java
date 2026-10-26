package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import java.util.Objects;

/** Workbook-level protection operations. */
public final class ExcelWorkbookProtection {
  private final ExcelWorkbook workbook;

  ExcelWorkbookProtection(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Enables workbook-level protection and password hashes with authoritative settings. */
  public ExcelWorkbook setWorkbookProtection(ResolvedWorkbookProtection protection) {
    return new ExcelSheetStateController().setWorkbookProtection(workbook, protection);
  }

  /** Clears workbook-level protection and password hashes entirely. */
  public ExcelWorkbook clearWorkbookProtection() {
    return new ExcelSheetStateController().clearWorkbookProtection(workbook);
  }
}
