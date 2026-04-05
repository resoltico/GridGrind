package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.protocol.dto.NamedRangeTarget;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeTarget record construction and engine conversion. */
class NamedRangeTargetTest {
  @Test
  void preservesProtocolRangeTextUntilEngineConversion() {
    NamedRangeTarget target = new NamedRangeTarget("Budget", "B4:A1");

    assertEquals("B4:A1", target.range());
    assertEquals(
        new ExcelNamedRangeTarget("Budget", "A1:B4"),
        WorkbookCommandConverter.toExcelNamedRangeTarget(target));
  }

  @Test
  void convertsNamedRangeTargetToEngineType() {
    NamedRangeTarget target = new NamedRangeTarget("Budget", "B4:C5");

    assertEquals(
        new ExcelNamedRangeTarget("Budget", "B4:C5"),
        WorkbookCommandConverter.toExcelNamedRangeTarget(target));
  }

  @Test
  void validatesNamedRangeTargetInputs() {
    assertThrows(NullPointerException.class, () -> new NamedRangeTarget(null, "A1"));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeTarget(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new NamedRangeTarget("ThisSheetNameIsFarTooLongForExcel", "A1"));
    assertThrows(NullPointerException.class, () -> new NamedRangeTarget("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeTarget("Budget", " "));
  }
}
