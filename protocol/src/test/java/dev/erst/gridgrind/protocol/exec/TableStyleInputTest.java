package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.protocol.dto.TableStyleInput;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing table style inputs and conversion. */
class TableStyleInputTest {
  @Test
  void convertsSupportedStyleVariants() {
    assertEquals(
        new ExcelTableStyle.None(),
        DefaultGridGrindRequestExecutor.toExcelTableStyle(new TableStyleInput.None()));
    assertEquals(
        new ExcelTableStyle.Named("TableStyleMedium2", true, false, true, false),
        DefaultGridGrindRequestExecutor.toExcelTableStyle(
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
