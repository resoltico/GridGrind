package dev.erst.gridgrind.core.model.drawing;

import static dev.erst.gridgrind.core.model.drawing.ShapeSpec.requireNonBlank;
import static dev.erst.gridgrind.core.model.drawing.ShapeSpec.requirePresetGeometryToken;

import dev.erst.gridgrind.core.vocabulary.ExcelAuthoredDrawingShapeKind;
import java.util.Objects;
import java.util.Optional;

/** Materialized shapes sharing source-free authored connectors. */
public sealed interface ResolvedShape permits ResolvedShape.SimpleShape, ShapeSpec.Connector {
  /** Returns the user-facing name for the authored drawing shape. */
  String name();

  /** Returns the required two-cell anchor describing the authored shape placement. */
  DrawingAnchorSpec anchor();

  /** Returns the authored shape kind represented by this definition. */
  ExcelAuthoredDrawingShapeKind kind();

  /** Simple shape with required preset geometry and optional text. */
  record SimpleShape(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
      String presetGeometryToken,
      Optional<String> text)
      implements ResolvedShape {
    public SimpleShape {
      name = requireNonBlank(name, "name");
      Objects.requireNonNull(anchor, "anchor must not be null");
      Objects.requireNonNull(text, "text must not be null");
      presetGeometryToken = requirePresetGeometryToken(presetGeometryToken);
      text.ifPresent(
          value -> {
            if (value.isBlank()) {
              throw new IllegalArgumentException("text must not be blank");
            }
          });
    }

    @Override
    public ExcelAuthoredDrawingShapeKind kind() {
      return ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE;
    }
  }
}
