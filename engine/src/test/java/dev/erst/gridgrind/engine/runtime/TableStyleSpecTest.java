package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing table style inputs and conversion. */
class TableStyleSpecTest {
  @Test
  void convertsSupportedStyleVariants() {
    assertEquals(
        new ExcelTableStyle.None(),
        WorkbookCommandTabularInputConverter.toExcelTableStyle(new TableStyleSpec.None()));
    assertEquals(
        new ExcelTableStyle.Named("TableStyleMedium2", true, false, true, false),
        WorkbookCommandTabularInputConverter.toExcelTableStyle(
            new TableStyleSpec.Named("TableStyleMedium2", true, false, true, false)));
  }

  @Test
  void validatesNamedStyleIdentifiers() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableStyleSpec.Named(" ", false, false, true, false));
    assertThrows(
        NullPointerException.class,
        () -> new TableStyleSpec.Named(null, false, false, true, false));
  }
}
