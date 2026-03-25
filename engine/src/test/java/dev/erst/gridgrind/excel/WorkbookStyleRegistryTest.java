package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for WorkbookStyleRegistry static utilities. */
class WorkbookStyleRegistryTest {
  @Test
  void resolveNumberFormat_returnsGeneralForNullOrBlank() {
    assertEquals("General", WorkbookStyleRegistry.resolveNumberFormat(null));
    assertEquals("General", WorkbookStyleRegistry.resolveNumberFormat(""));
    assertEquals("General", WorkbookStyleRegistry.resolveNumberFormat("   "));
  }

  @Test
  void resolveNumberFormat_returnsFormatStringWhenPopulated() {
    assertEquals("#,##0.00", WorkbookStyleRegistry.resolveNumberFormat("#,##0.00"));
    assertEquals("yyyy-mm-dd", WorkbookStyleRegistry.resolveNumberFormat("yyyy-mm-dd"));
  }
}
