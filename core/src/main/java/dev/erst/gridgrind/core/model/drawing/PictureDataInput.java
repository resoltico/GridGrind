package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.source.BinarySourceInput;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Objects;

/** Binary picture payload used for authored drawing pictures and embedded-object previews. */
public record PictureDataInput(ExcelPictureFormat format, BinarySourceInput source) {
  public PictureDataInput {
    Objects.requireNonNull(format, "format must not be null");
    Objects.requireNonNull(source, "source must not be null");
  }
}
