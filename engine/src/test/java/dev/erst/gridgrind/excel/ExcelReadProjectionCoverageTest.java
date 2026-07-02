package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Set;
import org.junit.jupiter.api.Test;

/** Coverage tests for internal read projection and window validation helpers. */
class ExcelReadProjectionCoverageTest {
  @Test
  void projectionDefaultsIncludesAndValidationCoverResidualBranches() {
    ExcelCellReadProjection defaults = ExcelCellReadProjection.defaults();
    ExcelCellReadProjection projection =
        new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.STYLE, ExcelCellReadFacet.VALUE));

    assertEquals(Set.of(ExcelCellReadFacet.VALUE), defaults.facets());
    assertTrue(defaults.includes(ExcelCellReadFacet.VALUE));
    assertFalse(defaults.includes(ExcelCellReadFacet.STYLE));
    assertTrue(projection.includes(ExcelCellReadFacet.STYLE));

    assertThrows(NullPointerException.class, () -> defaults.includes(null));
    assertThrows(NullPointerException.class, () -> new ExcelCellReadProjection(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCellReadProjection(Set.of()));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelCellReadProjection(java.util.Collections.singleton(null)));
  }

  @Test
  void readWindowRejectsNullTopLeftAddress() {
    ExcelReadWindow window = new ExcelReadWindow("B3", 2, 3);

    assertEquals("B3", window.topLeftAddress());
    assertThrows(NullPointerException.class, () -> new ExcelReadWindow(null, 1, 1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelReadWindow(" ", 1, 1));
  }
}
