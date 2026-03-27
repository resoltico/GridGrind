package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelNamedRangeScope record construction. */
class ExcelNamedRangeScopeTest {
  @Test
  void validatesSheetScopeInput() {
    assertDoesNotThrow(() -> new ExcelNamedRangeScope.WorkbookScope());
    assertDoesNotThrow(() -> new ExcelNamedRangeScope.SheetScope("Budget"));
    assertThrows(NullPointerException.class, () -> new ExcelNamedRangeScope.SheetScope(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelNamedRangeScope.SheetScope(" "));
  }
}
