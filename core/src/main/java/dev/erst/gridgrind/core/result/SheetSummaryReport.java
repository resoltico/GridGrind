package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import java.util.Objects;

/** Structural summary facts for one sheet. */
public record SheetSummaryReport(
    String sheetName,
    ExcelSheetVisibility visibility,
    SheetProtectionReport protection,
    int physicalRowCount,
    int lastRowIndex,
    int lastColumnIndex) {
  public SheetSummaryReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    Objects.requireNonNull(visibility, "visibility must not be null");
    Objects.requireNonNull(protection, "protection must not be null");
  }
}
