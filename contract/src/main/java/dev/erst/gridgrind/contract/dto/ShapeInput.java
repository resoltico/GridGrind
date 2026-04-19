package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelAuthoredDrawingShapeKind;
import java.util.Objects;

/** Authoritative simple-shape or connector creation or replacement payload. */
public record ShapeInput(
    String name,
    ExcelAuthoredDrawingShapeKind kind,
    DrawingAnchorInput anchor,
    String presetGeometryToken,
    TextSourceInput text) {
  public ShapeInput {
    name = requireNonBlank(name, "name");
    Objects.requireNonNull(kind, "kind must not be null");
    anchor = requireTwoCellAnchor(anchor);
    if (presetGeometryToken != null) {
      presetGeometryToken = presetGeometryToken.trim();
    }
    presetGeometryToken = normalizePresetGeometryToken(kind, presetGeometryToken);
    validateConnectorConfiguration(kind, presetGeometryToken, text);
    if (text instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
      throw new IllegalArgumentException("text must not be blank");
    }
  }

  private static String normalizePresetGeometryToken(
      ExcelAuthoredDrawingShapeKind kind, String presetGeometryToken) {
    if (kind == ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE
        && (presetGeometryToken == null || presetGeometryToken.isBlank())) {
      return "rect";
    }
    return presetGeometryToken;
  }

  private static void validateConnectorConfiguration(
      ExcelAuthoredDrawingShapeKind kind, String presetGeometryToken, TextSourceInput text) {
    if (kind != ExcelAuthoredDrawingShapeKind.CONNECTOR) {
      return;
    }
    if (presetGeometryToken != null) {
      throw new IllegalArgumentException("presetGeometryToken is only supported for SIMPLE_SHAPE");
    }
    if (text != null) {
      throw new IllegalArgumentException("text is only supported for SIMPLE_SHAPE");
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
