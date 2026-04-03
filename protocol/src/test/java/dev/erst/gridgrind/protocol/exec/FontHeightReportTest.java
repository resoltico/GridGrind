package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.protocol.dto.FontHeightReport;
import java.math.BigDecimal;
import org.junit.jupiter.api.Test;

/** Tests for analyzed font-height reports returned to protocol callers. */
class FontHeightReportTest {
  @Test
  void createsConsistentReportsFromEngineFontHeights() {
    FontHeightReport report =
        DefaultGridGrindRequestExecutor.toFontHeightReport(new ExcelFontHeight(230));

    assertEquals(230, report.twips());
    assertEquals(new BigDecimal("11.5"), report.points());
  }

  @Test
  void rejectsMismatchedPointAndTwipsValues() {
    assertThrows(NullPointerException.class, () -> new FontHeightReport(220, null));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightReport(0, BigDecimal.ZERO));
    assertThrows(
        IllegalArgumentException.class, () -> new FontHeightReport(220, new BigDecimal("11.5")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FontHeightReport(Short.MAX_VALUE + 1, new BigDecimal("1638.4")));
  }

  @Test
  void normalizesEquivalentPointValues() {
    FontHeightReport report = new FontHeightReport(220, new BigDecimal("11.0"));

    assertEquals(new BigDecimal("11"), report.points());
  }
}
