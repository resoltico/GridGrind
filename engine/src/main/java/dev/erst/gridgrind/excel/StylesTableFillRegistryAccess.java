package dev.erst.gridgrind.excel;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Objects;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;

/** Reflective access wrapper for POI's private styles-table fill registry. */
final class StylesTableFillRegistryAccess {
  private final Field fillsField;

  StylesTableFillRegistryAccess(Field fillsField) {
    this.fillsField = Objects.requireNonNull(fillsField, "fillsField must not be null");
  }

  static StylesTableFillRegistryAccess reflective() {
    return new StylesTableFillRegistryAccess(requiredField("fills"));
  }

  @SuppressWarnings("unchecked")
  List<XSSFCellFill> fills(StylesTable stylesTable) {
    try {
      return (List<XSSFCellFill>) fillsField.get(stylesTable);
    } catch (ReflectiveOperationException exception) {
      throw new IllegalStateException("Failed to access workbook fill registry", exception);
    }
  }

  @SuppressWarnings("PMD.AvoidAccessibilityAlteration")
  static Field requiredField(String fieldName) {
    try {
      Field field = StylesTable.class.getDeclaredField(fieldName);
      field.setAccessible(true);
      return field;
    } catch (ReflectiveOperationException exception) {
      throw new ExceptionInInitializerError(exception);
    }
  }
}
