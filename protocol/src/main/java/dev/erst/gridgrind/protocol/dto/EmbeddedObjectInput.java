package dev.erst.gridgrind.protocol.dto;

import java.util.Base64;
import java.util.Objects;

/** Authoritative embedded-object creation or replacement payload. */
public record EmbeddedObjectInput(
    String name,
    String label,
    String fileName,
    String command,
    String base64Data,
    PictureDataInput previewImage,
    DrawingAnchorInput anchor) {
  public EmbeddedObjectInput {
    requireNonBlank(name, "name");
    requireNonBlank(label, "label");
    requireNonBlank(fileName, "fileName");
    requireNonBlank(command, "command");
    requireBase64(base64Data, "base64Data");
    Objects.requireNonNull(previewImage, "previewImage must not be null");
    requireTwoCellAnchor(anchor);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static String requireBase64(String value, String fieldName) {
    String validatedValue = requireNonBlank(value, fieldName);
    try {
      Base64.getDecoder().decode(validatedValue);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(fieldName + " must be valid base64", exception);
    }
    return validatedValue;
  }

  private static DrawingAnchorInput requireTwoCellAnchor(DrawingAnchorInput anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell -> twoCell;
    };
  }
}
