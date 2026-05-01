package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import org.junit.jupiter.api.Test;

/** Tests for sheet copy-position protocol records. */
class SheetCopyPositionTest {
  @Test
  void appendAtEndConvertsToWorkbookCorePosition() {
    SheetCopyPosition.AppendAtEnd position = new SheetCopyPosition.AppendAtEnd();

    assertInstanceOf(
        ExcelSheetCopyPosition.AppendAtEnd.class,
        WorkbookCommandConverter.toExcelSheetCopyPosition(position));
  }

  @Test
  void atIndexValidatesAndConvertsToWorkbookCorePosition() {
    SheetCopyPosition.AtIndex position = new SheetCopyPosition.AtIndex(2);

    ExcelSheetCopyPosition.AtIndex converted =
        assertInstanceOf(
            ExcelSheetCopyPosition.AtIndex.class,
            WorkbookCommandConverter.toExcelSheetCopyPosition(position));

    assertEquals(2, converted.targetIndex());
    assertThrows(IllegalArgumentException.class, () -> new SheetCopyPosition.AtIndex(-1));
  }
}
