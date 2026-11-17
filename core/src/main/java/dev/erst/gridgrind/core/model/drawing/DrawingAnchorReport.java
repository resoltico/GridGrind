package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.FieldRule;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.ValueInvariantException;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.util.Objects;
import java.util.Optional;

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
      DrawingMarkerReport from,
      @FieldRule(kind = FieldRule.Kind.NUMBER_RANGE, minimum = 0, maximum = 27273042316900L)
          long widthEmu,
      @FieldRule(kind = FieldRule.Kind.NUMBER_RANGE, minimum = 0, maximum = 27273042316900L)
          long heightEmu,
      ExcelDrawingAnchorBehavior behavior)
      implements DrawingAnchorReport {
    public OneCell {
      Objects.requireNonNull(from, "from must not be null");
      behavior = Objects.requireNonNullElse(behavior, ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      if (widthEmu < 0L || widthEmu > 27273042316900L) {
        throw new ValueInvariantException(
            new MessageInvariant(
                "widthEmu must be between 0 and 27273042316900", Optional.of("widthEmu")));
      }
      if (heightEmu < 0L || heightEmu > 27273042316900L) {
        throw new ValueInvariantException(
            new MessageInvariant(
                "heightEmu must be between 0 and 27273042316900", Optional.of("heightEmu")));
      }
    }
  }

  /** Factual absolute anchor. */
  @Variant(id = "ABSOLUTE", summary = "Drawing anchor with stored absolute coordinates and size.")
  record Absolute(
      DrawingCoordinateReport x,
      DrawingCoordinateReport y,
      @FieldRule(kind = FieldRule.Kind.NUMBER_RANGE, minimum = 0, maximum = 27273042316900L)
          long widthEmu,
      @FieldRule(kind = FieldRule.Kind.NUMBER_RANGE, minimum = 0, maximum = 27273042316900L)
          long heightEmu,
      ExcelDrawingAnchorBehavior behavior)
      implements DrawingAnchorReport {
    public Absolute {
      behavior =
          Objects.requireNonNullElse(behavior, ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE);
      Objects.requireNonNull(x, "x must not be null");
      Objects.requireNonNull(y, "y must not be null");
      if (widthEmu < 0L || widthEmu > 27273042316900L) {
        throw new ValueInvariantException(
            new MessageInvariant(
                "widthEmu must be between 0 and 27273042316900", Optional.of("widthEmu")));
      }
      if (heightEmu < 0L || heightEmu > 27273042316900L) {
        throw new ValueInvariantException(
            new MessageInvariant(
                "heightEmu must be between 0 and 27273042316900", Optional.of("heightEmu")));
      }
    }
  }
}
