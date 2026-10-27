package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeScopeReport record construction. */
class ExcelNamedRangeScopeTest {
  @Test
  void validatesSheetScopeInput() {
    assertDoesNotThrow(() -> new NamedRangeScopeReport.Workbook());
    assertDoesNotThrow(
        () ->
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")));
    assertThrows(
        NullPointerException.class,
        () ->
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(" ")));
  }
}
