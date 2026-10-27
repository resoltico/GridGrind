package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.junit.jupiter.api.Test;

/** Tests for GridGrind-owned sheet visibility conversions. */
class ExcelSheetVisibilityTest {
  @Test
  void fromPoiMapsEverySupportedVisibility() {
    assertEquals(
        ExcelSheetVisibility.VISIBLE,
        SheetCollectionAdapter.fromPoiVisibility(SheetVisibility.VISIBLE));
    assertEquals(
        ExcelSheetVisibility.HIDDEN,
        SheetCollectionAdapter.fromPoiVisibility(SheetVisibility.HIDDEN));
    assertEquals(
        ExcelSheetVisibility.VERY_HIDDEN,
        SheetCollectionAdapter.fromPoiVisibility(SheetVisibility.VERY_HIDDEN));
    assertThrows(NullPointerException.class, () -> SheetCollectionAdapter.fromPoiVisibility(null));
  }

  @Test
  void toPoiMapsEverySupportedVisibility() {
    assertEquals(
        SheetVisibility.VISIBLE,
        SheetCollectionAdapter.toPoiVisibility(ExcelSheetVisibility.VISIBLE));
    assertEquals(
        SheetVisibility.HIDDEN,
        SheetCollectionAdapter.toPoiVisibility(ExcelSheetVisibility.HIDDEN));
    assertEquals(
        SheetVisibility.VERY_HIDDEN,
        SheetCollectionAdapter.toPoiVisibility(ExcelSheetVisibility.VERY_HIDDEN));
  }
}
