package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Authoritative embedded-object creation or replacement payload. */
@CatalogType(
    id = "EmbeddedObjectSpec",
    summary =
        "Named embedded-object authoring payload for SET_EMBEDDED_OBJECT. base64Data holds the embedded package bytes and previewImage holds the visible preview raster.",
    displayOrder = 39)
public record EmbeddedObjectSpec(
    String name,
    String label,
    String fileName,
    String command,
    BinarySourceSpec payload,
    PictureDataSpec previewImage,
    DrawingAnchorSpec anchor) {
  public EmbeddedObjectSpec {
    TextValidation.requireNonBlank(name, "name");
    TextValidation.requireNonBlank(label, "label");
    TextValidation.requireNonBlank(fileName, "fileName");
    TextValidation.requireNonBlank(command, "command");
    Objects.requireNonNull(payload, "payload must not be null");
    Objects.requireNonNull(previewImage, "previewImage must not be null");
    requireTwoCellAnchor(anchor);
  }

  private static DrawingAnchorSpec requireTwoCellAnchor(DrawingAnchorSpec anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorSpec.TwoCell twoCell -> twoCell;
    };
  }
}
