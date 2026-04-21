package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Covers same-sheet area normalization branches used by named ranges and chart references. */
class ExcelDrawingNormalizeAreaFormulaCoverageTest {
  @Test
  void normalizeAreaFormulaForPoiOnlyRewritesRepeatedSameSheetAreas() {
    assertEquals(
        "'Data Sheet'!$A$1:$B$2",
        ExcelChartSourceSupport.normalizeAreaFormulaForPoi("'Data Sheet'!$A$1:'Data Sheet'!$B$2"));
    assertEquals("A1", ExcelChartSourceSupport.normalizeAreaFormulaForPoi("A1"));
    assertEquals(
        "Data!$A$1:Other!$B$2",
        ExcelChartSourceSupport.normalizeAreaFormulaForPoi("Data!$A$1:Other!$B$2"));
    assertEquals(
        "Data!$A$1:$B$2:$C$3",
        ExcelChartSourceSupport.normalizeAreaFormulaForPoi("Data!$A$1:$B$2:$C$3"));
    assertEquals(
        "not-a-cell:Data!$B$2",
        ExcelChartSourceSupport.normalizeAreaFormulaForPoi("not-a-cell:Data!$B$2"));
  }
}
