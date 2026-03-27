package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeScope record construction and engine conversion. */
class NamedRangeScopeTest {
  @Test
  void convertsWorkbookAndSheetScopes() {
    assertEquals(
        new ExcelNamedRangeScope.WorkbookScope(),
        new NamedRangeScope.Workbook().toExcelNamedRangeScope());
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        new NamedRangeScope.Sheet("Budget").toExcelNamedRangeScope());
  }

  @Test
  void validatesSheetScopeInput() {
    assertThrows(NullPointerException.class, () -> new NamedRangeScope.Sheet(null));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeScope.Sheet(" "));
  }
}
