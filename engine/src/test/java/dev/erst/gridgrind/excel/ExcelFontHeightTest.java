package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import org.junit.jupiter.api.Test;

/** Tests for exact Excel font-height modeling in twips and point units. */
class ExcelFontHeightTest {
  @Test
  void convertsBetweenPointsAndTwipsExactly() {
    ExcelFontHeight wholePoints = ExcelFontHeight.fromPoints(new BigDecimal("13"));
    ExcelFontHeight fractionalPoints = ExcelFontHeight.fromPoints(new BigDecimal("11.5"));

    assertEquals(260, wholePoints.twips());
    assertEquals(new BigDecimal("13"), wholePoints.points());
    assertEquals(230, fractionalPoints.twips());
    assertEquals(new BigDecimal("11.5"), fractionalPoints.points());
  }

  @Test
  void rejectsNonPositiveOrInexactValues() {
    assertThrows(IllegalArgumentException.class, () -> new ExcelFontHeight(0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFontHeight(Short.MAX_VALUE + 1));
    assertThrows(IllegalArgumentException.class, () -> ExcelFontHeight.fromPoints(BigDecimal.ZERO));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelFontHeight.fromPoints(new BigDecimal("11.333")));
  }
}
