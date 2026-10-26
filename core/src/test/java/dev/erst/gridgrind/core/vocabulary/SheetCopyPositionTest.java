package dev.erst.gridgrind.core.vocabulary;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for copied-sheet placement records. */
class SheetCopyPositionTest {
  @Test
  void createsSupportedPlacementVariants() {
    assertInstanceOf(SheetCopyPosition.AppendAtEnd.class, new SheetCopyPosition.AppendAtEnd());
    assertEquals(2, new SheetCopyPosition.AtIndex(2).targetIndex());
  }

  @Test
  void rejectsNegativeTargetIndexes() {
    IllegalArgumentException exception =
        assertThrows(IllegalArgumentException.class, () -> new SheetCopyPosition.AtIndex(-1));
    assertEquals("targetIndex must be >= 0", exception.getMessage());
  }
}
