package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.Objects;

/** Structural summary facts for one sheet. */
public record SheetSummaryReport(
    SheetLabel sheetName,
    ExcelSheetVisibility visibility,
    SheetProtectionReport protection,
    int physicalRowCount,
    int lastRowIndex,
    int lastColumnIndex) {
  public SheetSummaryReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(visibility, "visibility must not be null");
    Objects.requireNonNull(protection, "protection must not be null");
  }
}
