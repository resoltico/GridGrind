package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;

/** Shared drawing anchor helpers for shipped example workbook plans. */
final class ExampleDrawingAnchors {
  private ExampleDrawingAnchors() {}

  static DrawingAnchorInput.TwoCell anchor(int fromColumn, int fromRow, int toColumn, int toRow) {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(fromColumn, fromRow),
        new DrawingMarkerInput(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }
}
