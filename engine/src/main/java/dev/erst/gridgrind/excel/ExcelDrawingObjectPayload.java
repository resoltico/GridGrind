package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Extracted binary drawing-object payload surfaced through read operations. */
public sealed interface ExcelDrawingObjectPayload
    permits ExcelDrawingObjectPayload.Picture, ExcelDrawingObjectPayload.EmbeddedObject {

  /** Sheet-local drawing object name. */
  String name();

  /** MIME type associated with the extracted payload bytes. */
  String contentType();

  /** Extracted payload bytes. */
  ExcelBinaryData data();

  /** SHA-256 digest of the extracted payload bytes. */
  String sha256();

  /** Extracted picture payload. */
  record Picture(
      String name,
      ExcelPictureFormat format,
      String contentType,
      String fileName,
      String sha256,
      ExcelBinaryData data,
      String description)
      implements ExcelDrawingObjectPayload {
    public Picture {
      validateCommon(name, contentType, sha256, data);
      Objects.requireNonNull(format, "format must not be null");
      fileName = requireNonBlank(fileName, "fileName");
      if (description != null && description.isBlank()) {
        throw new IllegalArgumentException("description must not be blank");
      }
    }
  }

  /** Extracted embedded-object payload. */
  record EmbeddedObject(
      String name,
      ExcelEmbeddedObjectPackagingKind packagingKind,
      String contentType,
      String fileName,
      String sha256,
      ExcelBinaryData data,
      String label,
      String command)
      implements ExcelDrawingObjectPayload {
    public EmbeddedObject {
      validateCommon(name, contentType, sha256, data);
      Objects.requireNonNull(packagingKind, "packagingKind must not be null");
      if (fileName != null && fileName.isBlank()) {
        throw new IllegalArgumentException("fileName must not be blank");
      }
      if (label != null && label.isBlank()) {
        throw new IllegalArgumentException("label must not be blank");
      }
      if (command != null && command.isBlank()) {
        throw new IllegalArgumentException("command must not be blank");
      }
    }
  }

  private static void validateCommon(
      String name, String contentType, String sha256, ExcelBinaryData data) {
    requireNonBlank(name, "name");
    requireNonBlank(contentType, "contentType");
    requireNonBlank(sha256, "sha256");
    Objects.requireNonNull(data, "data must not be null");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
