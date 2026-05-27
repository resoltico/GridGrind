package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing table style inputs and conversion. */
class TableStyleInputTest {
  @Test
  void convertsSupportedStyleVariants() {
    assertEquals(
        new ExcelTableStyle.None(),
        WorkbookCommandTabularInputConverter.toExcelTableStyle(new TableStyleInput.None()));
    assertEquals(
        new ExcelTableStyle.Named("TableStyleMedium2", true, false, true, false),
        WorkbookCommandTabularInputConverter.toExcelTableStyle(
            new TableStyleInput.Named("TableStyleMedium2", true, false, true, false)));
  }

  @Test
  void validatesNamedStyleIdentifiers() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableStyleInput.Named(" ", false, false, true, false));
    assertThrows(
        NullPointerException.class,
        () -> new TableStyleInput.Named(null, false, false, true, false));
  }
}
