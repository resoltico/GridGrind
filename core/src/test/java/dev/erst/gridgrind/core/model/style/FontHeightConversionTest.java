package dev.erst.gridgrind.core.model.style;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import org.junit.jupiter.api.Test;

/** Tests for exact Excel font-height modeling in twips and point units. */
class FontHeightConversionTest {
  @Test
  void authoredUnitsHaveExactPointAndTwipQueries() {
    FontHeightSpec points = new FontHeightSpec.Points(new BigDecimal("11.50"));
    FontHeightSpec twips = new FontHeightSpec.Twips(230);
    assertEquals(230, points.twips());
    assertEquals(230, twips.twips());
    assertEquals(0, new BigDecimal("11.5").compareTo(points.points()));
    assertEquals(0, points.points().compareTo(twips.points()));
    assertEquals(new BigDecimal("11.50"), points.points());
  }

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
