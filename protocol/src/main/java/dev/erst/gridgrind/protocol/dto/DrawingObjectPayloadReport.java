package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import java.util.Base64;
import java.util.Objects;

/** Extracted binary drawing-object payload returned by payload reads. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = DrawingObjectPayloadReport.Picture.class, name = "PICTURE"),
  @JsonSubTypes.Type(
      value = DrawingObjectPayloadReport.EmbeddedObject.class,
      name = "EMBEDDED_OBJECT")
})
public sealed interface DrawingObjectPayloadReport
    permits DrawingObjectPayloadReport.Picture, DrawingObjectPayloadReport.EmbeddedObject {

  /** Sheet-local drawing object name. */
  String name();

  /** MIME type associated with the extracted payload bytes. */
  String contentType();

  /** Base64-encoded extracted payload bytes. */
  String base64Data();

  /** SHA-256 digest of the extracted payload bytes. */
  String sha256();

  /** Extracted picture payload. */
  record Picture(
      String name,
      ExcelPictureFormat format,
      String contentType,
      String fileName,
      String sha256,
      String base64Data,
      String description)
      implements DrawingObjectPayloadReport {
    public Picture {
      validateCommon(name, contentType, sha256, base64Data);
      Objects.requireNonNull(format, "format must not be null");
      requireNonBlank(fileName, "fileName");
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
      String base64Data,
      String label,
      String command)
      implements DrawingObjectPayloadReport {
    public EmbeddedObject {
      validateCommon(name, contentType, sha256, base64Data);
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
      String name, String contentType, String sha256, String base64Data) {
    requireNonBlank(name, "name");
    requireNonBlank(contentType, "contentType");
    requireNonBlank(sha256, "sha256");
    requireBase64(base64Data, "base64Data");
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
}
