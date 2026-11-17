package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.FieldRule;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.ValueInvariantException;
import java.util.Objects;
import java.util.Optional;

/** One stored drawing coordinate, retaining its native unit and lexical representation. */
@Union(discriminator = "type", catalog = true, displayOrder = 16)
public sealed interface DrawingCoordinateReport
    permits DrawingCoordinateReport.Emu,
        DrawingCoordinateReport.Measured,
        DrawingCoordinateReport.Pixels {
  /** Integer DrawingML coordinate in English metric units. */
  @Variant(id = "EMU", summary = "Stored integer DrawingML coordinate in English metric units.")
  record Emu(
      @FieldRule(
              kind = FieldRule.Kind.NUMBER_RANGE,
              minimum = -27273042329600L,
              maximum = 27273042316900L)
          long value)
      implements DrawingCoordinateReport {
    public Emu {
      if (value < -27273042329600L || value > 27273042316900L) {
        throw new ValueInvariantException(
            new MessageInvariant(
                "value must be between -27273042329600 and 27273042316900", Optional.of("value")));
      }
    }
  }

  /** Native unit-bearing DrawingML coordinate, preserved without conversion. */
  @Variant(
      id = "MEASURED",
      summary = "Stored DrawingML measurement with its exact native units and lexical form.")
  record Measured(
      @FieldRule(
              kind = FieldRule.Kind.STRING_PATTERN,
              text = "-?[0-9]+(\\.[0-9]+)?(mm|cm|in|pt|pc|pi)")
          String valueText)
      implements DrawingCoordinateReport {
    public Measured {
      Objects.requireNonNull(valueText, "valueText must not be null");
      if (!valueText.matches("-?[0-9]+(\\.[0-9]+)?(mm|cm|in|pt|pc|pi)")) {
        throw new ValueInvariantException(
            new MessageInvariant(
                "valueText must be a native DrawingML measurement", Optional.of("valueText")));
      }
    }
  }

  /** VML client-anchor offset in pixels, without conversion to DrawingML units. */
  @Variant(id = "PIXELS", summary = "Stored nonnegative VML client-anchor offset in pixels.")
  record Pixels(
      @FieldRule(kind = FieldRule.Kind.NUMBER_RANGE, minimum = 0, maximum = Integer.MAX_VALUE)
          int value)
      implements DrawingCoordinateReport {
    public Pixels {
      if (value < 0) {
        throw new ValueInvariantException(
            new MessageInvariant("value must not be negative", Optional.of("value")));
      }
    }
  }
}
