package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import java.util.Objects;
import java.util.Optional;

/** Authoritative simple-shape or connector creation or replacement payload. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ShapeInput.SimpleShape.class, name = "SIMPLE_SHAPE"),
  @JsonSubTypes.Type(value = ShapeInput.Connector.class, name = "CONNECTOR")
})
public sealed interface ShapeInput permits ShapeInput.SimpleShape, ShapeInput.Connector {
  /** Stable drawing object name that replaces or creates the target shape. */
  String name();

  /** Anchor describing where the shape is positioned on the sheet. */
  DrawingAnchorInput anchor();

  /** Concrete authored drawing-shape family represented by this payload. */
  ExcelAuthoredDrawingShapeKind kind();

  /** Simple shape with required preset geometry and optional text. */
  record SimpleShape(
      String name,
      DrawingAnchorInput anchor,
      String presetGeometryToken,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<TextSourceInput> text)
      implements ShapeInput {
    public SimpleShape {
      name = requireNonBlank(name, "name");
      anchor = requireTwoCellAnchor(anchor);
      Objects.requireNonNull(text, "text must not be null");
      presetGeometryToken = requirePresetGeometryToken(presetGeometryToken);
      text.ifPresent(
          value -> {
            if (value instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
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
  record Connector(String name, DrawingAnchorInput anchor) implements ShapeInput {
    public Connector {
      name = requireNonBlank(name, "name");
      anchor = requireTwoCellAnchor(anchor);
    }

    @Override
    public ExcelAuthoredDrawingShapeKind kind() {
      return ExcelAuthoredDrawingShapeKind.CONNECTOR;
    }
  }

  private static String requirePresetGeometryToken(String presetGeometryToken) {
    Objects.requireNonNull(presetGeometryToken, "presetGeometryToken must not be null");
    String normalized = presetGeometryToken.trim();
    if (normalized.isBlank()) {
      throw new IllegalArgumentException("presetGeometryToken must not be blank");
    }
    return normalized;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static DrawingAnchorInput requireTwoCellAnchor(DrawingAnchorInput anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell -> twoCell;
    };
  }
}
