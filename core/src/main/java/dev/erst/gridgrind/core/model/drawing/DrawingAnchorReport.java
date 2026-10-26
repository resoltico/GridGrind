package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.util.Objects;

/** Factual drawing anchor report returned by drawing-object reads. */
@Union(discriminator = "type", catalog = true, displayOrder = 15)
public sealed interface DrawingAnchorReport
    permits dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport.TwoCell,
        dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport.OneCell,
        dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport.Absolute {

  /** Factual two-cell anchor. */
  @Variant(id = "TWO_CELL", summary = "Drawing anchor spanning one start and end marker.")
  record TwoCell(
      DrawingMarkerReport from, DrawingMarkerReport to, ExcelDrawingAnchorBehavior behavior)
      implements DrawingAnchorReport {
    public TwoCell {
      Objects.requireNonNull(from, "from must not be null");
      Objects.requireNonNull(to, "to must not be null");
      behavior = Objects.requireNonNullElse(behavior, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    }
  }

  /** Factual one-cell anchor. */
  @Variant(id = "ONE_CELL", summary = "Drawing anchor with one start marker plus explicit size.")
  record OneCell(
      DrawingMarkerReport from, long widthEmu, long heightEmu, ExcelDrawingAnchorBehavior behavior)
      implements DrawingAnchorReport {
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

  /** Factual absolute anchor. */
  @Variant(id = "ABSOLUTE", summary = "Drawing anchor with absolute EMU coordinates and size.")
  record Absolute(
      long xEmu, long yEmu, long widthEmu, long heightEmu, ExcelDrawingAnchorBehavior behavior)
      implements DrawingAnchorReport {
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
