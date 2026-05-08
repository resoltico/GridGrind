package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import java.util.Objects;
import java.util.Optional;

/** Authoritative simple-shape or connector creation or replacement payload. */
public sealed interface ExcelShapeDefinition
    permits ExcelShapeDefinition.SimpleShape, ExcelShapeDefinition.Connector {
  /** Returns the user-facing name for the authored drawing shape. */
  String name();

  /** Returns the required two-cell anchor describing the authored shape placement. */
  ExcelDrawingAnchor.TwoCell anchor();

  /** Returns the authored shape kind represented by this definition. */
  ExcelAuthoredDrawingShapeKind kind();

  /** Simple shape with required preset geometry and optional text. */
  record SimpleShape(
      String name,
      ExcelDrawingAnchor.TwoCell anchor,
      String presetGeometryToken,
      Optional<String> text)
      implements ExcelShapeDefinition {
    public SimpleShape {
      Objects.requireNonNull(name, "name must not be null");
      if (name.isBlank()) {
        throw new IllegalArgumentException("name must not be blank");
      }
      Objects.requireNonNull(anchor, "anchor must not be null");
      Objects.requireNonNull(text, "text must not be null");
      Objects.requireNonNull(presetGeometryToken, "presetGeometryToken must not be null");
      presetGeometryToken = presetGeometryToken.trim();
      if (presetGeometryToken.isBlank()) {
        throw new IllegalArgumentException("presetGeometryToken must not be blank");
      }
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

  /** Connector drawing with no geometry token and no text payload. */
  record Connector(String name, ExcelDrawingAnchor.TwoCell anchor) implements ExcelShapeDefinition {
    public Connector {
      Objects.requireNonNull(name, "name must not be null");
      if (name.isBlank()) {
        throw new IllegalArgumentException("name must not be blank");
      }
      Objects.requireNonNull(anchor, "anchor must not be null");
    }

    @Override
    public ExcelAuthoredDrawingShapeKind kind() {
      return ExcelAuthoredDrawingShapeKind.CONNECTOR;
    }
  }
}
