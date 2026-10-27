package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.Optional;

/** Authoritative picture creation or replacement payload. */
@CatalogType(
    id = "PictureSpec",
    summary = "Named picture-authoring payload for SET_PICTURE.",
    displayOrder = 37)
public record PictureSpec(
    String name,
    PictureDataSpec image,
    DrawingAnchorSpec anchor,
    Optional<TextSourceSpec> description) {
  public PictureSpec {
    name = TextValidation.requireNonBlank(name, "name");
    Objects.requireNonNull(image, "image must not be null");
    anchor = requireTwoCellAnchor(anchor);
    Objects.requireNonNull(description, "description must not be null");
    description.ifPresent(
        value -> {
          if (value instanceof TextSourceSpec.Inline inline && inline.text().isBlank()) {
            throw new IllegalArgumentException("description must not be blank");
          }
        });
  }

  private static DrawingAnchorSpec requireTwoCellAnchor(DrawingAnchorSpec anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorSpec.TwoCell twoCell -> twoCell;
    };
  }
}
