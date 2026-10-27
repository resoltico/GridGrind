package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import java.util.List;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Coverage tests for internal read projection and window validation helpers. */
class ExcelReadProjectionCoverageTest {
  @Test
  void projectionDefaultsIncludesAndValidationCoverResidualBranches() {
    CellReadProjection defaults = CellReadProjection.defaults();
    CellReadProjection projection =
        new CellReadProjection(List.copyOf(Set.of(CellReadFacet.STYLE, CellReadFacet.VALUE)));

    assertEquals(List.of(CellReadFacet.VALUE), defaults.facets());
    assertTrue(defaults.includes(CellReadFacet.VALUE));
    assertFalse(defaults.includes(CellReadFacet.STYLE));
    assertTrue(projection.includes(CellReadFacet.STYLE));

    assertThrows(NullPointerException.class, () -> defaults.includes(null));
    assertThrows(NullPointerException.class, () -> new CellReadProjection(null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellReadProjection(List.copyOf(Set.of())));
    assertThrows(
        NullPointerException.class,
        () ->
            new CellReadProjection(
                new java.util.ArrayList<>(java.util.Arrays.asList((CellReadFacet) null))));
  }
}
