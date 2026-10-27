package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.engine.names.NamedRangeAdapter;
import org.junit.jupiter.api.Test;

/** Covers same-sheet area normalization branches used by named ranges and chart references. */
class ExcelDrawingNormalizeAreaFormulaCoverageTest {
  @Test
  void normalizeAreaFormulaForPoiOnlyRewritesRepeatedSameSheetAreas() {
    assertEquals(
        "'Data Sheet'!$A$1:$B$2",
        NamedRangeAdapter.normalizeAreaFormulaForPoi("'Data Sheet'!$A$1:'Data Sheet'!$B$2"));
    assertEquals("A1", NamedRangeAdapter.normalizeAreaFormulaForPoi("A1"));
    assertEquals(
        "Data!$A$1:Other!$B$2",
        NamedRangeAdapter.normalizeAreaFormulaForPoi("Data!$A$1:Other!$B$2"));
    assertEquals(
        "Data!$A$1:$B$2:$C$3", NamedRangeAdapter.normalizeAreaFormulaForPoi("Data!$A$1:$B$2:$C$3"));
    assertEquals(
        "not-a-cell:Data!$B$2",
        NamedRangeAdapter.normalizeAreaFormulaForPoi("not-a-cell:Data!$B$2"));
  }
}
