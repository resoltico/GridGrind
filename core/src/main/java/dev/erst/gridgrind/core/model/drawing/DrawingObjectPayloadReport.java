package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Base64;
import java.util.Objects;
import java.util.Optional;

/** Extracted binary drawing-object payload returned by payload reads. */
@Union(discriminator = "type")
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
  @Variant(id = "PICTURE", summary = "Extracted picture payload.")
  record Picture(
      String name,
      ExcelPictureFormat format,
      String contentType,
      String fileName,
      String sha256,
      String base64Data,
      Optional<String> description)
      implements DrawingObjectPayloadReport {
    public Picture {
      validateCommon(name, contentType, sha256, base64Data);
      Objects.requireNonNull(format, "format must not be null");
      TextValidation.requireNonBlank(fileName, "fileName");
      description = requireOptionalText(description, "description");
    }
  }

  /** Extracted embedded-object payload. */
  @Variant(id = "EMBEDDED_OBJECT", summary = "Extracted embedded-object payload.")
  record EmbeddedObject(
      String name,
      ExcelEmbeddedObjectPackagingKind packagingKind,
      String contentType,
      Optional<String> fileName,
      String sha256,
      String base64Data,
      Optional<String> label,
      Optional<String> command)
      implements DrawingObjectPayloadReport {
    public EmbeddedObject {
      validateCommon(name, contentType, sha256, base64Data);
      Objects.requireNonNull(packagingKind, "packagingKind must not be null");
      fileName = requireOptionalText(fileName, "fileName");
      label = requireOptionalText(label, "label");
      command = requireOptionalText(command, "command");
    }
  }

  private static void validateCommon(
      String name, String contentType, String sha256, String base64Data) {
    TextValidation.requireNonBlank(name, "name");
    TextValidation.requireNonBlank(contentType, "contentType");
    TextValidation.requireNonBlank(sha256, "sha256");
    requireBase64(base64Data, "base64Data");
  }

  private static Optional<String> requireOptionalText(Optional<String> value, String fieldName) {
    return Objects.requireNonNull(value, fieldName + " must not be null")
        .map(text -> TextValidation.requireNonBlank(text, fieldName));
  }

  private static String requireBase64(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    try {
      Base64.getDecoder().decode(value);
    } catch (IllegalArgumentException exception) {
      throw new IllegalArgumentException(fieldName + " must be valid base64", exception);
    }
    return value;
  }
}
