package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for copied-sheet placement records. */
class ExcelSheetCopyPositionTest {
  @Test
  void createsSupportedPlacementVariants() {
    assertInstanceOf(
        ExcelSheetCopyPosition.AppendAtEnd.class, new ExcelSheetCopyPosition.AppendAtEnd());
    assertEquals(2, new ExcelSheetCopyPosition.AtIndex(2).targetIndex());
  }

  @Test
  void rejectsNegativeTargetIndexes() {
    IllegalArgumentException exception =
        assertThrows(IllegalArgumentException.class, () -> new ExcelSheetCopyPosition.AtIndex(-1));
    assertEquals("targetIndex must not be negative", exception.getMessage());
  }
}
