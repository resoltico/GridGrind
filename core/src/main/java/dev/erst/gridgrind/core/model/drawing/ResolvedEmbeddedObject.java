package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Embedded-object payload with materialized package and preview content. */
public record ResolvedEmbeddedObject(
    String name,
    String label,
    String fileName,
    String command,
    BinaryContent payload,
    ExcelPictureFormat previewFormat,
    BinaryContent previewImage,
    DrawingAnchorSpec.TwoCell anchor) {
  public ResolvedEmbeddedObject {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(label, "label must not be null");
    Objects.requireNonNull(fileName, "fileName must not be null");
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(payload, "payload must not be null");
    Objects.requireNonNull(previewFormat, "previewFormat must not be null");
    Objects.requireNonNull(previewImage, "previewImage must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    TextValidation.requireNonBlank(name, "name");
    TextValidation.requireNonBlank(label, "label");
    TextValidation.requireNonBlank(fileName, "fileName");
    TextValidation.requireNonBlank(command, "command");
  }
}
