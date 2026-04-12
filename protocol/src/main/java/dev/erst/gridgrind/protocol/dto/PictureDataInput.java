package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelPictureFormat;
import java.util.Base64;
import java.util.Objects;

/** Binary picture payload used for authored drawing pictures and embedded-object previews. */
public record PictureDataInput(ExcelPictureFormat format, String base64Data) {
  public PictureDataInput {
    Objects.requireNonNull(format, "format must not be null");
    base64Data = requireBase64(base64Data, "base64Data");
  }

  private static String requireBase64(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    try {
      Base64.getDecoder().decode(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(fieldName + " must be valid base64", exception);
    }
    return value;
  }
}
