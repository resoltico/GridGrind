package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.apache.poi.ss.usermodel.SheetVisibility;
import org.junit.jupiter.api.Test;

/** Tests for GridGrind-owned sheet visibility conversions. */
class ExcelSheetVisibilityTest {
  @Test
  void fromPoiMapsEverySupportedVisibility() {
    assertEquals(
        ExcelSheetVisibility.VISIBLE, ExcelSheetVisibility.fromPoi(SheetVisibility.VISIBLE));
    assertEquals(ExcelSheetVisibility.HIDDEN, ExcelSheetVisibility.fromPoi(SheetVisibility.HIDDEN));
    assertEquals(
        ExcelSheetVisibility.VERY_HIDDEN,
        ExcelSheetVisibility.fromPoi(SheetVisibility.VERY_HIDDEN));
    assertThrows(NullPointerException.class, () -> ExcelSheetVisibility.fromPoi(null));
  }

  @Test
  void toPoiMapsEverySupportedVisibility() {
    assertEquals(SheetVisibility.VISIBLE, ExcelSheetVisibility.VISIBLE.toPoi());
    assertEquals(SheetVisibility.HIDDEN, ExcelSheetVisibility.HIDDEN.toPoi());
    assertEquals(SheetVisibility.VERY_HIDDEN, ExcelSheetVisibility.VERY_HIDDEN.toPoi());
  }
}
