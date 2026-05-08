package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Objects;
import java.util.Optional;

/** Authoritative picture creation or replacement payload. */
public record ExcelPictureDefinition(
    String name,
    ExcelBinaryData imageData,
    ExcelPictureFormat format,
    ExcelDrawingAnchor.TwoCell anchor,
    Optional<String> description) {
  public ExcelPictureDefinition {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(imageData, "imageData must not be null");
    Objects.requireNonNull(format, "format must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(description, "description must not be null");
    description.ifPresent(
        value -> {
          if (value.isBlank()) {
            throw new IllegalArgumentException("description must not be blank");
          }
        });
  }
}
