package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import java.util.Objects;

/** Binary picture payload used for authored drawing pictures and embedded-object previews. */
public record PictureDataInput(ExcelPictureFormat format, BinarySourceInput source) {
  public PictureDataInput {
    Objects.requireNonNull(format, "format must not be null");
    Objects.requireNonNull(source, "source must not be null");
  }
}
