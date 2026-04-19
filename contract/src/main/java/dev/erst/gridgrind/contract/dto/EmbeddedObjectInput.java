package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.BinarySourceInput;
import java.util.Objects;

/** Authoritative embedded-object creation or replacement payload. */
public record EmbeddedObjectInput(
    String name,
    String label,
    String fileName,
    String command,
    BinarySourceInput payload,
    PictureDataInput previewImage,
    DrawingAnchorInput anchor) {
  public EmbeddedObjectInput {
    requireNonBlank(name, "name");
    requireNonBlank(label, "label");
    requireNonBlank(fileName, "fileName");
    requireNonBlank(command, "command");
    Objects.requireNonNull(payload, "payload must not be null");
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

  private static DrawingAnchorInput requireTwoCellAnchor(DrawingAnchorInput anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell -> twoCell;
    };
  }
}
