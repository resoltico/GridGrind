package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.FieldRule;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.ValueInvariantException;
import java.util.Objects;
import java.util.Optional;

/** One factual cell-relative drawing marker returned by drawing reads. */
@CatalogType(
    id = "DrawingMarkerReport",
    summary = "Exact cell-relative drawing marker report.",
    displayOrder = 96)
public record DrawingMarkerReport(
    @FieldRule(kind = FieldRule.Kind.NUMBER_RANGE, minimum = 0, maximum = Integer.MAX_VALUE)
        int columnIndex,
    @FieldRule(kind = FieldRule.Kind.NUMBER_RANGE, minimum = 0, maximum = Integer.MAX_VALUE)
        int rowIndex,
    DrawingCoordinateReport dx,
    DrawingCoordinateReport dy) {
  public DrawingMarkerReport {
    if (columnIndex < 0) {
      throw new ValueInvariantException(
          new MessageInvariant("columnIndex must not be negative", Optional.of("columnIndex")));
    }
    if (rowIndex < 0) {
      throw new ValueInvariantException(
          new MessageInvariant("rowIndex must not be negative", Optional.of("rowIndex")));
    }
    Objects.requireNonNull(dx, "dx must not be null");
    Objects.requireNonNull(dy, "dy must not be null");
  }
}
