package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;

/** Shared drawing anchor helpers for shipped example workbook plans. */
final class ExampleDrawingAnchors {
  private ExampleDrawingAnchors() {}

  static DrawingAnchorSpec.TwoCell anchor(int fromColumn, int fromRow, int toColumn, int toRow) {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(fromColumn, fromRow),
        new DrawingMarkerSpec(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }
}
