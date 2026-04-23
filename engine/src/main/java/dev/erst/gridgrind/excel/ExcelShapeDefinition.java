package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import java.util.Objects;

/** Authoritative simple-shape or connector creation or replacement payload. */
public record ExcelShapeDefinition(
    String name,
    ExcelAuthoredDrawingShapeKind kind,
    ExcelDrawingAnchor.TwoCell anchor,
    String presetGeometryToken,
    String text) {
  public ExcelShapeDefinition {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(kind, "kind must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    if (presetGeometryToken != null) {
      presetGeometryToken = presetGeometryToken.trim();
    }
    if (kind == ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE) {
      if (presetGeometryToken == null || presetGeometryToken.isBlank()) {
        presetGeometryToken = "rect";
      }
    } else {
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
}
