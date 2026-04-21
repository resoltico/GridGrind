package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.invoke.MethodHandles;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STGradientType;

/** Tests the supported fill-registry seam used by the workbook style registry. */
class StylesTableFillRegistryAccessTest {
  @Test
  void poiApiReturnsWorkbookFillRegistryAndAppendsDistinctGradientFills() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StylesTableFillRegistryAccess access = StylesTableFillRegistryAccess.poiApi();
      var stylesTable = workbook.getStylesSource();
      var fills = access.fills(stylesTable);

      assertNotNull(fills);
      assertFalse(fills.isEmpty());
      int baselineSize = fills.size();

      CTFill linearFill = CTFill.Factory.newInstance();
      linearFill.addNewGradientFill().setDegree(42.5d);
      CTFill pathFill = CTFill.Factory.newInstance();
      var pathGradient = pathFill.addNewGradientFill();
      pathGradient.setType(STGradientType.PATH);
      pathGradient.setLeft(0.1d);
      pathGradient.setRight(0.2d);
      pathGradient.setTop(0.3d);
      pathGradient.setBottom(0.4d);

      XSSFCellFill linear = new XSSFCellFill(linearFill, stylesTable.getIndexedColors());
      XSSFCellFill path = new XSSFCellFill(pathFill, stylesTable.getIndexedColors());

      assertTrue(linear.equals(path), "POI equality ignores gradient geometry");

      int linearId = access.appendFill(stylesTable, linear);
      int pathId = access.appendFill(stylesTable, path);

      assertEquals(baselineSize, linearId);
      assertEquals(baselineSize + 1, pathId);
      assertEquals(baselineSize + 2, access.fills(stylesTable).size());
      assertEquals(linearFill.xmlText(), stylesTable.getFillAt(linearId).getCTFill().xmlText());
      assertEquals(pathFill.xmlText(), stylesTable.getFillAt(pathId).getCTFill().xmlText());
    }
  }

  @Test
  void fillsRejectNullStylesTable() {
    assertThrows(
        NullPointerException.class, () -> StylesTableFillRegistryAccess.poiApi().fills(null));
  }

  @Test
  void appendFillRejectsNullArguments() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      StylesTableFillRegistryAccess access = StylesTableFillRegistryAccess.poiApi();

      assertThrows(NullPointerException.class, () -> access.appendFill(null, null));
      assertThrows(
          NullPointerException.class, () -> access.appendFill(workbook.getStylesSource(), null));
    }
  }

  @Test
  void requireFillsFieldRejectsLookupsWithoutPrivateStylesTableAccess() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> StylesTableFillRegistryAccess.requireFillsField(MethodHandles.publicLookup()));

    assertInstanceOf(IllegalAccessException.class, failure.getCause());
  }
}
