package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.BinarySourceInput;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Objects;

/** Binary picture payload used for authored drawing pictures and embedded-object previews. */
@CatalogType(
    id = "PictureDataInput",
    summary = "One picture payload with explicit format and base64-encoded binary data.",
    displayOrder = 36)
public record PictureDataInput(ExcelPictureFormat format, BinarySourceInput source) {
  public PictureDataInput {
    Objects.requireNonNull(format, "format must not be null");
    Objects.requireNonNull(source, "source must not be null");
  }
}
