package dev.erst.gridgrind.core.result;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import org.junit.jupiter.api.Test;

/** Preserves the summary bounds formerly enforced by native producer records. */
class SheetSummaryReportTest {
  @Test
  void acceptsEmptySheetSentinels() {
    var report = summary(0, -1, -1);
    assertEquals(0, report.physicalRowCount());
    assertEquals(-1, report.lastRowIndex());
    assertEquals(-1, report.lastColumnIndex());
  }

  @Test
  void rejectsImpossibleCountsAndIndexes() {
    assertThrows(IllegalArgumentException.class, () -> summary(-1, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> summary(0, -2, 0));
    assertThrows(IllegalArgumentException.class, () -> summary(0, 0, -2));
  }

  private static SheetSummaryReport summary(int count, int row, int column) {
    return new SheetSummaryReport(
        new SheetLabel("Observed"),
        ExcelSheetVisibility.VISIBLE,
        new SheetProtectionReport.Unprotected(),
        count,
        row,
        column);
  }
}
