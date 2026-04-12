package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.lang.reflect.Method;
import org.junit.jupiter.api.Test;

/** Covers same-sheet area normalization branches used by named ranges and chart references. */
@SuppressWarnings("PMD.AvoidAccessibilityAlteration")
class ExcelDrawingNormalizeAreaFormulaCoverageTest {
  @Test
  void normalizeAreaFormulaForPoiOnlyRewritesRepeatedSameSheetAreas()
      throws ReflectiveOperationException {
    ExcelDrawingController controller = new ExcelDrawingController();

    assertEquals(
        "'Data Sheet'!$A$1:$B$2",
        invokeNormalizeAreaFormulaForPoi(controller, "'Data Sheet'!$A$1:'Data Sheet'!$B$2"));
    assertEquals("A1", invokeNormalizeAreaFormulaForPoi(controller, "A1"));
    assertEquals(
        "Data!$A$1:Other!$B$2",
        invokeNormalizeAreaFormulaForPoi(controller, "Data!$A$1:Other!$B$2"));
    assertEquals(
        "Data!$A$1:$B$2:$C$3", invokeNormalizeAreaFormulaForPoi(controller, "Data!$A$1:$B$2:$C$3"));
    assertEquals(
        "not-a-cell:Data!$B$2",
        invokeNormalizeAreaFormulaForPoi(controller, "not-a-cell:Data!$B$2"));
  }

  private String invokeNormalizeAreaFormulaForPoi(ExcelDrawingController controller, String formula)
      throws ReflectiveOperationException {
    Method method =
        ExcelDrawingController.class.getDeclaredMethod("normalizeAreaFormulaForPoi", String.class);
    method.setAccessible(true);
    return (String) method.invoke(controller, formula);
  }
}
