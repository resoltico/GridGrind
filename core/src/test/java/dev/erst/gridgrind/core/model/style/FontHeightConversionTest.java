package dev.erst.gridgrind.core.model.style;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import org.junit.jupiter.api.Test;

/** Tests for exact Excel font-height modeling in twips and point units. */
class FontHeightConversionTest {
  @Test
  void convertsBetweenPointsAndTwipsExactly() {
    FontHeightReport wholePoints = FontHeightReport.fromPoints(new BigDecimal("13"));
    FontHeightReport fractionalPoints = FontHeightReport.fromPoints(new BigDecimal("11.5"));

    assertEquals(260, wholePoints.twips());
    assertEquals(new BigDecimal("13"), wholePoints.points());
    assertEquals(230, fractionalPoints.twips());
    assertEquals(new BigDecimal("11.5"), fractionalPoints.points());
  }

  @Test
  void rejectsNonPositiveOrInexactValues() {
    assertThrows(IllegalArgumentException.class, () -> FontHeightReport.fromTwips(0));
    assertThrows(
        IllegalArgumentException.class, () -> FontHeightReport.fromTwips(Short.MAX_VALUE + 1));
    assertThrows(
        IllegalArgumentException.class, () -> FontHeightReport.fromPoints(BigDecimal.ZERO));
    assertThrows(
        IllegalArgumentException.class,
        () -> FontHeightReport.fromPoints(new BigDecimal("11.333")));
  }
}
