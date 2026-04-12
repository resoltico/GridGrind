package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import java.util.Objects;

/** Authoritative simple-shape or connector creation or replacement payload. */
public record ShapeInput(
    String name,
    ExcelAuthoredDrawingShapeKind kind,
    DrawingAnchorInput anchor,
    String presetGeometryToken,
    String text) {
  public ShapeInput {
    name = requireNonBlank(name, "name");
    Objects.requireNonNull(kind, "kind must not be null");
    anchor = requireTwoCellAnchor(anchor);
    if (presetGeometryToken != null) {
      presetGeometryToken = presetGeometryToken.trim();
    }
    if (kind == ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE) {
      if (presetGeometryToken == null || presetGeometryToken.isBlank()) {
        presetGeometryToken = "rect";
      }
    } else if (kind == ExcelAuthoredDrawingShapeKind.CONNECTOR) {
      if (presetGeometryToken != null) {
        throw new IllegalArgumentException(
            "presetGeometryToken is only supported for SIMPLE_SHAPE");
      }
      if (text != null) {
        throw new IllegalArgumentException("text is only supported for SIMPLE_SHAPE");
      }
    }
    if (text != null && text.isBlank()) {
      throw new IllegalArgumentException("text must not be blank");
    }
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
