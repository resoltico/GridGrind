package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import org.junit.jupiter.api.Test;

/** Tests for protocol font-height input variants and engine conversion. */
class FontHeightInputTest {
  @Test
  void convertsPointsAndTwipsToEngineFontHeight() {
    FontHeightInput.Points points = new FontHeightInput.Points(new BigDecimal("11.5"));
    FontHeightInput.Twips twips = new FontHeightInput.Twips(260);

    assertEquals(230, points.toExcelFontHeight().twips());
    assertEquals(new BigDecimal("11.5"), points.toExcelFontHeight().points());
    assertEquals(new BigDecimal("13"), twips.toExcelFontHeight().points());
  }

  @Test
  void validatesProtocolFontHeightVariants() {
    assertThrows(NullPointerException.class, () -> new FontHeightInput.Points(null));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightInput.Points(BigDecimal.ZERO));
    assertThrows(
        IllegalArgumentException.class, () -> new FontHeightInput.Points(new BigDecimal("11.333")));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightInput.Twips(0));
  }
}
