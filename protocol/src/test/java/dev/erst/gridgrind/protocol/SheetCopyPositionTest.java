package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import org.junit.jupiter.api.Test;

/** Tests for sheet copy-position protocol records. */
class SheetCopyPositionTest {
  @Test
  void appendAtEndConvertsToWorkbookCorePosition() {
    SheetCopyPosition.AppendAtEnd position = new SheetCopyPosition.AppendAtEnd();

    assertInstanceOf(ExcelSheetCopyPosition.AppendAtEnd.class, position.toExcelSheetCopyPosition());
  }

  @Test
  void atIndexValidatesAndConvertsToWorkbookCorePosition() {
    SheetCopyPosition.AtIndex position = new SheetCopyPosition.AtIndex(2);

    ExcelSheetCopyPosition.AtIndex converted =
        assertInstanceOf(ExcelSheetCopyPosition.AtIndex.class, position.toExcelSheetCopyPosition());

    assertEquals(2, converted.targetIndex());
    assertThrows(NullPointerException.class, () -> new SheetCopyPosition.AtIndex(null));
    assertThrows(IllegalArgumentException.class, () -> new SheetCopyPosition.AtIndex(-1));
  }
}
