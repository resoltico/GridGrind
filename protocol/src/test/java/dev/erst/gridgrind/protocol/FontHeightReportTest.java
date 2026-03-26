package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelFontHeight;
import java.math.BigDecimal;
import org.junit.jupiter.api.Test;

/** Tests for analyzed font-height reports returned to protocol callers. */
class FontHeightReportTest {
  @Test
  void createsConsistentReportsFromEngineFontHeights() {
    FontHeightReport report = FontHeightReport.fromExcelFontHeight(new ExcelFontHeight(230));

    assertEquals(230, report.twips());
    assertEquals(new BigDecimal("11.5"), report.points());
  }

  @Test
  void rejectsMismatchedPointAndTwipsValues() {
    assertThrows(NullPointerException.class, () -> new FontHeightReport(220, null));
    assertThrows(
        IllegalArgumentException.class, () -> new FontHeightReport(220, new BigDecimal("11.5")));
  }
}
