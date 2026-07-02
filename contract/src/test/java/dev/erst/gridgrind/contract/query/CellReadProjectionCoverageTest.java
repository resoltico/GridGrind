package dev.erst.gridgrind.contract.query;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage tests for projection defaulting, deduplication, and validation branches. */
class CellReadProjectionCoverageTest {
  @Test
  void defaultsFactoryIncludesAndValidationCoverProjectionBranches() {
    CellReadProjection defaults = CellReadProjection.defaults();
    CellReadProjection projection =
        CellReadProjection.of(
            CellReadFacet.STYLE, CellReadFacet.VALUE, CellReadFacet.STYLE, CellReadFacet.FORMAT);

    assertEquals(List.of(CellReadFacet.VALUE), defaults.facets());
    assertTrue(defaults.includes(CellReadFacet.VALUE));
    assertFalse(defaults.includes(CellReadFacet.STYLE));
    assertThrows(NullPointerException.class, () -> defaults.includes(null));

    assertEquals(
        List.of(CellReadFacet.STYLE, CellReadFacet.VALUE, CellReadFacet.FORMAT),
        projection.facets());

    assertThrows(NullPointerException.class, () -> new CellReadProjection(null));
    assertThrows(IllegalArgumentException.class, () -> new CellReadProjection(List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new CellReadProjection(java.util.Arrays.asList((CellReadFacet) null)));
    assertThrows(NullPointerException.class, () -> CellReadProjection.of(null));
    assertThrows(
        NullPointerException.class,
        () -> CellReadProjection.of(CellReadFacet.VALUE, (CellReadFacet[]) null));
    assertThrows(
        NullPointerException.class,
        () -> CellReadProjection.of(CellReadFacet.VALUE, CellReadFacet.STYLE, null));
  }
}
