package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Factual drawing-object report returned by drawing reads. */
@Union(discriminator = "type")
public sealed interface DrawingObjectReport
    permits DrawingObjectReport.Picture,
        DrawingObjectReport.Chart,
        DrawingObjectReport.Shape,
        DrawingObjectReport.EmbeddedObject,
        DrawingObjectReport.SignatureLine {

  /** Sheet-local drawing object name. */
  String name();

  /** Stored anchor backing the drawing object. */
  DrawingAnchorReport anchor();

  /** Factual picture report. */
  @Variant(id = "PICTURE", summary = "Factual picture report.")
  record Picture(
      String name,
      DrawingAnchorReport anchor,
      ExcelPictureFormat format,
      String contentType,
      long byteSize,
      String sha256,
      Optional<Integer> widthPixels,
      Optional<Integer> heightPixels,
      Optional<String> description)
      implements DrawingObjectReport {
    public Picture {
      Objects.requireNonNull(widthPixels, "widthPixels must not be null");
      Objects.requireNonNull(heightPixels, "heightPixels must not be null");
      validateCommon(name, anchor);
      Objects.requireNonNull(format, "format must not be null");
      contentType = TextValidation.requireNonBlank(contentType, "contentType");
      sha256 = TextValidation.requireNonBlank(sha256, "sha256");
      if (byteSize < 0L) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      if (widthPixels.filter(value -> value < 0).isPresent()) {
        throw new IllegalArgumentException("widthPixels must not be negative");
      }
      if (heightPixels.filter(value -> value < 0).isPresent()) {
        throw new IllegalArgumentException("heightPixels must not be negative");
      }
      description = requireOptionalText(description, "description");
    }
  }

  /** Factual chart report surfaced through the drawing-object inventory. */
  @Variant(
      id = "CHART",
      summary = "Factual chart report surfaced through the drawing-object inventory.")
  record Chart(
      String name,
      DrawingAnchorReport anchor,
      boolean supported,
      List<String> plotTypeTokens,
      String title)
      implements DrawingObjectReport {
    public Chart {
      validateCommon(name, anchor);
      plotTypeTokens = copyNonBlankStrings(plotTypeTokens, "plotTypeTokens");
      Objects.requireNonNull(title, "title must not be null");
    }
  }

  /** Factual shape report. */
  @Variant(id = "SHAPE", summary = "Factual shape report.")
  record Shape(
      String name,
      DrawingAnchorReport anchor,
      ExcelDrawingShapeKind kind,
      Optional<String> presetGeometryToken,
      Optional<String> text,
      int childCount)
      implements DrawingObjectReport {
    public Shape {
      validateCommon(name, anchor);
      Objects.requireNonNull(kind, "kind must not be null");
      presetGeometryToken = requireOptionalText(presetGeometryToken, "presetGeometryToken");
      text = requireOptionalText(text, "text");
      if (childCount < 0) {
        throw new IllegalArgumentException("childCount must not be negative");
      }
    }
  }

  /** Factual embedded-object report. */
  @Variant(id = "EMBEDDED_OBJECT", summary = "Factual embedded-object report.")
  record EmbeddedObject(
      String name,
      DrawingAnchorReport anchor,
      ExcelEmbeddedObjectPackagingKind packagingKind,
      Optional<String> label,
      Optional<String> fileName,
      Optional<String> command,
      String contentType,
      long byteSize,
      String sha256,
      Optional<ExcelPictureFormat> previewFormat,
      Optional<Long> previewByteSize,
      Optional<String> previewSha256)
      implements DrawingObjectReport {
    public EmbeddedObject {
      Objects.requireNonNull(previewFormat, "previewFormat must not be null");
      Objects.requireNonNull(previewByteSize, "previewByteSize must not be null");
      validateCommon(name, anchor);
      Objects.requireNonNull(packagingKind, "packagingKind must not be null");
      contentType = TextValidation.requireNonBlank(contentType, "contentType");
      sha256 = TextValidation.requireNonBlank(sha256, "sha256");
      label = requireOptionalText(label, "label");
      fileName = requireOptionalText(fileName, "fileName");
      command = requireOptionalText(command, "command");
      if (byteSize < 0L) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      if (previewFormat.isEmpty() && previewByteSize.isPresent()) {
        throw new IllegalArgumentException("previewByteSize requires previewFormat");
      }
      previewSha256 = requireOptionalText(previewSha256, "previewSha256");
      if (previewFormat.isEmpty() && previewSha256.isPresent()) {
        throw new IllegalArgumentException("previewSha256 requires previewFormat");
      }
      if (previewByteSize.filter(value -> value < 0L).isPresent()) {
        throw new IllegalArgumentException("previewByteSize must not be negative");
      }
    }
  }

  /** Factual signature-line report. */
  @Variant(id = "SIGNATURE_LINE", summary = "Factual signature-line report.")
  record SignatureLine(
      String name,
      DrawingAnchorReport anchor,
      Optional<SignatureSetup> setup,
      Optional<SignaturePreview> preview)
      implements DrawingObjectReport {
    public SignatureLine {
      validateCommon(name, anchor);
      setup = requireOptional(setup, "setup");
      preview = requireOptional(preview, "preview");
    }
  }

  /** Optional authored signature-line setup metadata. */
  record SignatureSetup(
      Optional<String> setupId,
      Optional<Boolean> allowComments,
      Optional<String> signingInstructions,
      Optional<String> suggestedSigner,
      Optional<String> suggestedSigner2,
      Optional<String> suggestedSignerEmail) {
    public SignatureSetup {
      setupId = requireOptionalText(setupId, "setupId");
      allowComments = requireOptional(allowComments, "allowComments");
      signingInstructions = requireOptionalText(signingInstructions, "signingInstructions");
      suggestedSigner = requireOptionalText(suggestedSigner, "suggestedSigner");
      suggestedSigner2 = requireOptionalText(suggestedSigner2, "suggestedSigner2");
      suggestedSignerEmail = requireOptionalText(suggestedSignerEmail, "suggestedSignerEmail");
    }
  }

  /** Optional preview-image metadata surfaced for a signature line. */
  record SignaturePreview(
      ExcelPictureFormat format,
      String contentType,
      long byteSize,
      Optional<String> sha256,
      Optional<Integer> widthPixels,
      Optional<Integer> heightPixels) {
    public SignaturePreview {
      Objects.requireNonNull(format, "format must not be null");
      contentType = TextValidation.requireNonBlank(contentType, "contentType");
      sha256 = requireOptionalText(sha256, "previewSha256");
      if (byteSize < 0L) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      widthPixels = requireOptional(widthPixels, "widthPixels");
      heightPixels = requireOptional(heightPixels, "heightPixels");
      widthPixels.ifPresent(
          width -> {
            if (width < 0) {
              throw new IllegalArgumentException("previewWidthPixels must not be negative");
            }
          });
      heightPixels.ifPresent(
          height -> {
            if (height < 0) {
              throw new IllegalArgumentException("previewHeightPixels must not be negative");
            }
          });
    }
  }

  private static void validateCommon(String name, DrawingAnchorReport anchor) {
    TextValidation.requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
  }

  private static Optional<String> requireOptionalText(Optional<String> value, String fieldName) {
    Optional<String> normalized = requireOptional(value, fieldName);
    return normalized.map(text -> TextValidation.requireNonBlank(text, fieldName));
  }

  private static List<String> copyNonBlankStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = new ArrayList<>(values.size());
    for (String value : values) {
      copy.add(TextValidation.requireNonBlank(value, fieldName + " value"));
    }
    return List.copyOf(copy);
  }

  private static <T> Optional<T> requireOptional(Optional<T> value, String fieldName) {
    return Objects.requireNonNull(value, fieldName + " must not be null");
  }
}
