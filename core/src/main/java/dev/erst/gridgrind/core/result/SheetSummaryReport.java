package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.Objects;

/** Structural summary facts for one sheet. */
@CatalogType(
    id = "SheetSummaryReport",
    summary = "Exact sheet summary report including visibility, protection, and structural counts.",
    displayOrder = 77)
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
    if (physicalRowCount < 0)
      throw new IllegalArgumentException("physicalRowCount must not be negative");
    if (lastRowIndex < -1)
      throw new IllegalArgumentException(
          "lastRowIndex "
              + lastRowIndex
              + " must be greater than or equal to -1; empty sheets report -1");
    if (lastColumnIndex < -1)
      throw new IllegalArgumentException(
          "lastColumnIndex "
              + lastColumnIndex
              + " must be greater than or equal to -1; empty sheets report -1");
  }
}
