package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.junit.jupiter.api.Test;

/** Tests for GridGrind-owned sheet visibility conversions. */
class ExcelSheetVisibilityTest {
  @Test
  void fromPoiMapsEverySupportedVisibility() {
    assertEquals(
        ExcelSheetVisibility.VISIBLE,
        ExcelSheetVisibilityPoiBridge.fromPoi(SheetVisibility.VISIBLE));
    assertEquals(
        ExcelSheetVisibility.HIDDEN, ExcelSheetVisibilityPoiBridge.fromPoi(SheetVisibility.HIDDEN));
    assertEquals(
        ExcelSheetVisibility.VERY_HIDDEN,
        ExcelSheetVisibilityPoiBridge.fromPoi(SheetVisibility.VERY_HIDDEN));
    assertThrows(NullPointerException.class, () -> ExcelSheetVisibilityPoiBridge.fromPoi(null));
  }

  @Test
  void toPoiMapsEverySupportedVisibility() {
    assertEquals(
        SheetVisibility.VISIBLE, ExcelSheetVisibilityPoiBridge.toPoi(ExcelSheetVisibility.VISIBLE));
    assertEquals(
        SheetVisibility.HIDDEN, ExcelSheetVisibilityPoiBridge.toPoi(ExcelSheetVisibility.HIDDEN));
    assertEquals(
        SheetVisibility.VERY_HIDDEN,
        ExcelSheetVisibilityPoiBridge.toPoi(ExcelSheetVisibility.VERY_HIDDEN));
  }
}
