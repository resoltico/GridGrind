package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** Authoritative picture creation or replacement payload. */
public record PictureInput(
    String name, PictureDataInput image, DrawingAnchorInput anchor, String description) {
  public PictureInput {
    name = requireNonBlank(name, "name");
    Objects.requireNonNull(image, "image must not be null");
    anchor = requireTwoCellAnchor(anchor);
    if (description != null && description.isBlank()) {
      throw new IllegalArgumentException("description must not be blank");
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
