package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeTarget record construction and engine conversion. */
class NamedRangeTargetTest {
  @Test
  void preservesProtocolRangeTextUntilEngineConversion() {
    NamedRangeTarget.Range target = NamedRangeTarget.range("Budget", "B4:A1");

    assertEquals("B4:A1", target.range());
    assertEquals(
        ExcelNamedRangeTarget.range("Budget", "A1:B4"),
        WorkbookCommandConverter.toExcelNamedRangeTarget(target));
  }

  @Test
  void convertsNamedRangeTargetToEngineType() {
    NamedRangeTarget.Range target = NamedRangeTarget.range("Budget", "B4:C5");

    assertEquals(
        ExcelNamedRangeTarget.range("Budget", "B4:C5"),
        WorkbookCommandConverter.toExcelNamedRangeTarget(target));
  }

  @Test
  void validatesNamedRangeTargetInputs() {
    assertThrows(NullPointerException.class, () -> NamedRangeTarget.range(null, "A1"));
    assertThrows(IllegalArgumentException.class, () -> NamedRangeTarget.range(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.range("ThisSheetNameIsFarTooLongForExcel", "A1"));
    assertThrows(NullPointerException.class, () -> NamedRangeTarget.range("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> NamedRangeTarget.range("Budget", " "));
  }
}
