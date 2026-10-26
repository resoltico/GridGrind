package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Objects;
import java.util.Optional;

/** Picture payload with materialized binary content and canonical authored geometry. */
public record ResolvedPicture(
    String name,
    BinaryContent imageData,
    ExcelPictureFormat format,
    DrawingAnchorSpec.TwoCell anchor,
    Optional<String> description) {
  public ResolvedPicture {
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
