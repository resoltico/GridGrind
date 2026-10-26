package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.BinarySourceInput;
import java.util.Objects;

/** Authoritative embedded-object creation or replacement payload. */
@CatalogType(
    id = "EmbeddedObjectInput",
    summary =
        "Named embedded-object authoring payload for SET_EMBEDDED_OBJECT. base64Data holds the embedded package bytes and previewImage holds the visible preview raster.",
    displayOrder = 39)
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
