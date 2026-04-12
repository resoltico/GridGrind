package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Authoritative picture creation or replacement payload. */
public record ExcelPictureDefinition(
    String name,
    ExcelBinaryData imageData,
    ExcelPictureFormat format,
    ExcelDrawingAnchor.TwoCell anchor,
    String description) {
  public ExcelPictureDefinition {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(imageData, "imageData must not be null");
    Objects.requireNonNull(format, "format must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    if (description != null && description.isBlank()) {
      throw new IllegalArgumentException("description must not be blank");
    }
  }
}
