package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import java.math.BigDecimal;
import org.junit.jupiter.api.Test;

/** Tests for protocol font-height input variants and engine conversion. */
class FontHeightSpecTest {
  @Test
  void convertsPointsAndTwipsToEngineFontHeight() {
    FontHeightSpec.Points points = new FontHeightSpec.Points(new BigDecimal("11.5"));
    FontHeightSpec.Twips twips = new FontHeightSpec.Twips(260);

    assertEquals(
        230, WorkbookCommandCellValueSpecConverter.toExcelFontHeight(points).orElseThrow().twips());
    assertEquals(
        new BigDecimal("11.5"),
        WorkbookCommandCellValueSpecConverter.toExcelFontHeight(points).orElseThrow().points());
    assertEquals(
        new BigDecimal("13"),
        WorkbookCommandCellValueSpecConverter.toExcelFontHeight(twips).orElseThrow().points());
  }

  @Test
  void validatesProtocolFontHeightVariants() {
    assertThrows(NullPointerException.class, () -> new FontHeightSpec.Points(null));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightSpec.Points(BigDecimal.ZERO));
    assertThrows(
        IllegalArgumentException.class, () -> new FontHeightSpec.Points(new BigDecimal("11.333")));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightSpec.Twips(0));
    assertThrows(
        IllegalArgumentException.class, () -> new FontHeightSpec.Twips(Short.MAX_VALUE + 1));
  }
}
