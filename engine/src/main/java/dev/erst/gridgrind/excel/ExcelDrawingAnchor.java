package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import java.util.Objects;

/** Immutable factual or authored drawing anchor state. */
public sealed interface ExcelDrawingAnchor
    permits ExcelDrawingAnchor.TwoCell, ExcelDrawingAnchor.OneCell, ExcelDrawingAnchor.Absolute {

  /** Two-cell anchor backed by a start and end marker. */
  record TwoCell(
      ExcelDrawingMarker from, ExcelDrawingMarker to, ExcelDrawingAnchorBehavior behavior)
      implements ExcelDrawingAnchor {
    public TwoCell {
      Objects.requireNonNull(from, "from must not be null");
      Objects.requireNonNull(to, "to must not be null");
      behavior = Objects.requireNonNullElse(behavior, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
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

  /** One-cell anchor backed by a start marker plus an absolute size. */
  record OneCell(
      ExcelDrawingMarker from, long widthEmu, long heightEmu, ExcelDrawingAnchorBehavior behavior)
      implements ExcelDrawingAnchor {
    public OneCell {
      Objects.requireNonNull(from, "from must not be null");
      behavior = Objects.requireNonNullElse(behavior, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      if (widthEmu <= 0L) {
        throw new IllegalArgumentException("widthEmu must be greater than 0");
      }
      if (heightEmu <= 0L) {
        throw new IllegalArgumentException("heightEmu must be greater than 0");
      }
    }
  }

  /** Absolute anchor backed by a fixed position and size in EMUs. */
  record Absolute(
      long xEmu, long yEmu, long widthEmu, long heightEmu, ExcelDrawingAnchorBehavior behavior)
      implements ExcelDrawingAnchor {
    public Absolute {
      behavior =
          Objects.requireNonNullElse(behavior, ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE);
      if (xEmu < 0L) {
        throw new IllegalArgumentException("xEmu must not be negative");
      }
      if (yEmu < 0L) {
        throw new IllegalArgumentException("yEmu must not be negative");
      }
      if (widthEmu <= 0L) {
        throw new IllegalArgumentException("widthEmu must be greater than 0");
      }
      if (heightEmu <= 0L) {
        throw new IllegalArgumentException("heightEmu must be greater than 0");
      }
    }
  }
}
