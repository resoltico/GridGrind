package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelAuthoredDrawingShapeKind;
import java.util.Objects;
import java.util.Optional;

/** Authoritative simple-shape or connector creation or replacement payload. */
@Union(discriminator = "type", catalog = true, displayOrder = 24)
public sealed interface ShapeSpec
    permits dev.erst.gridgrind.core.model.drawing.ShapeSpec.SimpleShape,
        dev.erst.gridgrind.core.model.drawing.ShapeSpec.Connector {
  /** Stable drawing object name that replaces or creates the target shape. */
  String name();

  /** Anchor describing where the shape is positioned on the sheet. */
  DrawingAnchorSpec anchor();

  /** Concrete authored drawing-shape family represented by this payload. */
  ExcelAuthoredDrawingShapeKind kind();

  /** Simple shape with required preset geometry and optional text. */
  @Variant(
      id = "SIMPLE_SHAPE",
      summary =
          "Write one authored simple shape with required geometry token and optional text payload.")
  record SimpleShape(
      String name,
      DrawingAnchorSpec anchor,
      String presetGeometryToken,
      Optional<TextSourceSpec> text)
      implements ShapeSpec {
    public SimpleShape {
      name = requireNonBlank(name, "name");
      anchor = requireTwoCellAnchor(anchor);
      Objects.requireNonNull(text, "text must not be null");
      presetGeometryToken = requirePresetGeometryToken(presetGeometryToken);
      text.ifPresent(
          value -> {
            if (value instanceof TextSourceSpec.Inline inline && inline.text().isBlank()) {
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
  @Variant(
      id = "CONNECTOR",
      summary = "Write one authored connector with no geometry token or text payload.")
  record Connector(String name, DrawingAnchorSpec anchor) implements ShapeSpec, ResolvedShape {
    public Connector {
      name = requireNonBlank(name, "name");
      anchor = requireTwoCellAnchor(anchor);
    }

    @Override
    public ExcelAuthoredDrawingShapeKind kind() {
      return ExcelAuthoredDrawingShapeKind.CONNECTOR;
    }
  }

  /** Normalizes a required geometry token for authored and materialized simple shapes. */
  static String requirePresetGeometryToken(String presetGeometryToken) {
    Objects.requireNonNull(presetGeometryToken, "presetGeometryToken must not be null");
    String normalized = presetGeometryToken.trim();
    if (normalized.isBlank()) {
      throw new IllegalArgumentException("presetGeometryToken must not be blank");
    }
    return normalized;
  }

  /** Validates required shape text shared by authored and materialized payloads. */
  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static DrawingAnchorSpec requireTwoCellAnchor(DrawingAnchorSpec anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorSpec.TwoCell twoCell -> twoCell;
    };
  }
}
