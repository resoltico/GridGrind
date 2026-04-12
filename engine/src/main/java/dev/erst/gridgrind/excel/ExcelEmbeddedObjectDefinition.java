package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Authoritative embedded-object creation or replacement payload. */
public record ExcelEmbeddedObjectDefinition(
    String name,
    String label,
    String fileName,
    String command,
    ExcelBinaryData payload,
    ExcelPictureFormat previewFormat,
    ExcelBinaryData previewImage,
    ExcelDrawingAnchor.TwoCell anchor) {
  public ExcelEmbeddedObjectDefinition {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(label, "label must not be null");
    Objects.requireNonNull(fileName, "fileName must not be null");
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(payload, "payload must not be null");
    Objects.requireNonNull(previewFormat, "previewFormat must not be null");
    Objects.requireNonNull(previewImage, "previewImage must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    if (label.isBlank()) {
      throw new IllegalArgumentException("label must not be blank");
    }
    if (fileName.isBlank()) {
      throw new IllegalArgumentException("fileName must not be blank");
    }
    if (command.isBlank()) {
      throw new IllegalArgumentException("command must not be blank");
    }
  }
}
