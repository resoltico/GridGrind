package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.protocol.dto.NamedRangeScope;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeScope record construction and engine conversion. */
class NamedRangeScopeTest {
  @Test
  void convertsWorkbookAndSheetScopes() {
    assertEquals(
        new ExcelNamedRangeScope.WorkbookScope(),
        DefaultGridGrindRequestExecutor.toExcelNamedRangeScope(new NamedRangeScope.Workbook()));
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        DefaultGridGrindRequestExecutor.toExcelNamedRangeScope(
            new NamedRangeScope.Sheet("Budget")));
  }

  @Test
  void validatesSheetScopeInput() {
    assertThrows(NullPointerException.class, () -> new NamedRangeScope.Sheet(null));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeScope.Sheet(" "));
  }
}
