package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import java.util.Objects;

/** Factual drawing anchor report returned by drawing-object reads. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = DrawingAnchorReport.TwoCell.class, name = "TWO_CELL"),
  @JsonSubTypes.Type(value = DrawingAnchorReport.OneCell.class, name = "ONE_CELL"),
  @JsonSubTypes.Type(value = DrawingAnchorReport.Absolute.class, name = "ABSOLUTE")
})
public sealed interface DrawingAnchorReport
    permits DrawingAnchorReport.TwoCell, DrawingAnchorReport.OneCell, DrawingAnchorReport.Absolute {

  /** Factual two-cell anchor. */
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
