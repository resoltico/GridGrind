package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeTarget record construction and engine conversion. */
class NamedRangeTargetTest {
  @Test
  void preservesProtocolRangeTextUntilEngineConversion() {
    NamedRangeTarget.Range target =
        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4:A1"));

    assertEquals("B4:A1", target.range().value());
    assertEquals(
        ExcelNamedRangeTarget.range("Budget", "A1:B4"),
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeTarget(target));
  }

  @Test
  void convertsNamedRangeTargetToEngineType() {
    NamedRangeTarget.Range target =
        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4:C5"));

    assertEquals(
        ExcelNamedRangeTarget.range("Budget", "B4:C5"),
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeTarget(target));
  }

  @Test
  void validatesNamedRangeTargetInputs() {
    assertThrows(
        NullPointerException.class, () -> NamedRangeTarget.range(null, new RangeRef("A1")));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.range(new SheetName(" "), new RangeRef("A1")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            NamedRangeTarget.range(
                new SheetName("ThisSheetNameIsFarTooLongForExcel"), new RangeRef("A1")));
    assertThrows(
        NullPointerException.class, () -> NamedRangeTarget.range(new SheetName("Budget"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.range(new SheetName("Budget"), new RangeRef(" ")));
  }
}
