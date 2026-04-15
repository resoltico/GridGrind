package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.reflect.Field;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests the reflective seam used to access POI's private fill registry. */
class StylesTableFillRegistryAccessTest {
  @Test
  void reflectiveAccessReturnsWorkbookFillRegistry() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var fills = StylesTableFillRegistryAccess.reflective().fills(workbook.getStylesSource());

      assertNotNull(fills);
      assertFalse(fills.isEmpty());
    }
  }

  @Test
  void requiredFieldFailsFastWhenPoiStructureChanges() {
    assertThrows(
        ExceptionInInitializerError.class,
        () -> StylesTableFillRegistryAccess.requiredField("missingFillsField"));
  }

  @Test
  void fillsFailsWhenFieldIsNotAccessible() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Field fillsField = StylesTable.class.getDeclaredField("fills");
      fillsField.setAccessible(false);

      IllegalStateException exception =
          assertThrows(
              IllegalStateException.class,
              () ->
                  new StylesTableFillRegistryAccess(fillsField).fills(workbook.getStylesSource()));

      assertTrue(exception.getCause() instanceof IllegalAccessException);
    }
  }
}
