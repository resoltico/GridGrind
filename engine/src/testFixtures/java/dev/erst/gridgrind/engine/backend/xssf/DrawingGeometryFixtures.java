package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingCoordinateReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;

/** Builds expected factual geometry from authored fixture coordinates, independently of POI. */
public final class DrawingGeometryFixtures {
  private DrawingGeometryFixtures() {}

  /** Builds the expected two-cell facts for one authored fixture rectangle. */
  public static DrawingAnchorReport.TwoCell report(DrawingAnchorSpec.TwoCell anchor) {
    return new DrawingAnchorReport.TwoCell(
        marker(anchor.from()), marker(anchor.to()), anchor.behavior());
  }

  /** Copies one authored fixture's exact coordinates into its factual representation. */
  public static DrawingMarkerReport marker(DrawingMarkerSpec marker) {
    return new DrawingMarkerReport(
        marker.columnIndex(),
        marker.rowIndex(),
        new DrawingCoordinateReport.Emu(marker.dx()),
        new DrawingCoordinateReport.Emu(marker.dy()));
  }

  /** Builds VML client-anchor facts in their stored pixel units. */
  public static DrawingAnchorReport.TwoCell signatureReport(DrawingAnchorSpec.TwoCell anchor) {
    return new DrawingAnchorReport.TwoCell(
        pixelMarker(anchor.from()), pixelMarker(anchor.to()), anchor.behavior());
  }

  private static DrawingMarkerReport pixelMarker(DrawingMarkerSpec marker) {
    return new DrawingMarkerReport(
        marker.columnIndex(),
        marker.rowIndex(),
        new DrawingCoordinateReport.Pixels(marker.dx()),
        new DrawingCoordinateReport.Pixels(marker.dy()));
  }
}
