package dev.erst.gridgrind.core.model.drawing;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.problem.ValueInvariantException;
import org.junit.jupiter.api.Test;

/** Exact observed native facets are independent of writable rectangle policy. */
class DrawingCoordinateReportTest {
  @Test
  void coordinateKindsKeepUnitsAndExactNativeBounds() {
    assertEquals(-27273042329600L, new DrawingCoordinateReport.Emu(-27273042329600L).value());
    assertEquals(27273042316900L, new DrawingCoordinateReport.Emu(27273042316900L).value());
    assertThrows(
        ValueInvariantException.class, () -> new DrawingCoordinateReport.Emu(-27273042329601L));
    assertThrows(
        ValueInvariantException.class, () -> new DrawingCoordinateReport.Emu(27273042316901L));
    assertEquals("-01.00cm", new DrawingCoordinateReport.Measured("-01.00cm").valueText());
    assertThrows(ValueInvariantException.class, () -> new DrawingCoordinateReport.Measured("1zz"));
    assertThrows(ValueInvariantException.class, () -> new DrawingCoordinateReport.Measured(""));
    assertThrows(NullPointerException.class, () -> new DrawingCoordinateReport.Measured(null));
    assertEquals(17, new DrawingCoordinateReport.Pixels(17).value());
    assertThrows(ValueInvariantException.class, () -> new DrawingCoordinateReport.Pixels(-1));
  }

  @Test
  void observedIndicesAndExtentsKeepNativeDomain() {
    var marker =
        new DrawingMarkerReport(
            Integer.MAX_VALUE,
            Integer.MAX_VALUE,
            new DrawingCoordinateReport.Emu(-1),
            new DrawingCoordinateReport.Measured("1cm"));
    var anchor = new DrawingAnchorReport.OneCell(marker, 0, 27273042316900L, null);
    assertEquals(0, anchor.widthEmu());
    assertThrows(
        ValueInvariantException.class, () -> new DrawingAnchorReport.OneCell(marker, -1, 0, null));
    assertThrows(
        ValueInvariantException.class,
        () -> new DrawingAnchorReport.OneCell(marker, 0, 27273042316901L, null));
    assertThrows(
        NullPointerException.class,
        () -> new DrawingMarkerReport(0, 0, null, new DrawingCoordinateReport.Emu(0)));
  }
}
