package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.util.Objects;

/** Authored drawing anchor input. Phase 5 currently supports two-cell anchors for mutation. */
@Union(discriminator = "type", catalog = true, displayOrder = 48)
public sealed interface DrawingAnchorInput
    permits dev.erst.gridgrind.core.model.drawing.DrawingAnchorInput.TwoCell {

  /** Two-cell authored anchor with explicit from and to markers. */
  @Variant(
      id = "TWO_CELL",
      summary =
          "Authored two-cell drawing anchor with explicit start and end markers. behavior is explicit on the wire; the Java authoring surface exposes DrawingAnchorInput.TwoCell.moveAndResize(...).")
  record TwoCell(
      DrawingMarkerInput from, DrawingMarkerInput to, ExcelDrawingAnchorBehavior behavior)
      implements DrawingAnchorInput {
    /** Creates one two-cell anchor that moves and resizes with the authored rectangle. */
    public static TwoCell moveAndResize(DrawingMarkerInput from, DrawingMarkerInput to) {
      return new TwoCell(from, to, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    }

    public TwoCell {
      Objects.requireNonNull(from, "from must not be null");
      Objects.requireNonNull(to, "to must not be null");
      Objects.requireNonNull(behavior, "behavior must not be null");
      if (to.rowIndex() < from.rowIndex()) {
        throw new IllegalArgumentException("to.rowIndex must not be less than from.rowIndex");
      }
      if (to.columnIndex() < from.columnIndex()) {
        throw new IllegalArgumentException("to.columnIndex must not be less than from.columnIndex");
      }
      if (to.rowIndex() == from.rowIndex() && to.dy() < from.dy()) {
        throw new IllegalArgumentException("to.dy must not be less than from.dy on the same row");
      }
      if (to.columnIndex() == from.columnIndex() && to.dx() < from.dx()) {
        throw new IllegalArgumentException(
            "to.dx must not be less than from.dx on the same column");
      }
    }
  }
}
