package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Factual drawing-object report returned by drawing reads. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = DrawingObjectReport.Picture.class, name = "PICTURE"),
  @JsonSubTypes.Type(value = DrawingObjectReport.Chart.class, name = "CHART"),
  @JsonSubTypes.Type(value = DrawingObjectReport.Shape.class, name = "SHAPE"),
  @JsonSubTypes.Type(value = DrawingObjectReport.EmbeddedObject.class, name = "EMBEDDED_OBJECT"),
  @JsonSubTypes.Type(value = DrawingObjectReport.SignatureLine.class, name = "SIGNATURE_LINE")
})
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
  record Picture(
      String name,
      DrawingAnchorReport anchor,
      ExcelPictureFormat format,
      String contentType,
      long byteSize,
      String sha256,
      Integer widthPixels,
      Integer heightPixels,
      String description)
      implements DrawingObjectReport {
    public Picture {
      validateCommon(name, anchor);
      Objects.requireNonNull(format, "format must not be null");
      contentType = requireNonBlank(contentType, "contentType");
      sha256 = requireNonBlank(sha256, "sha256");
      if (byteSize < 0L) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      if (widthPixels != null && widthPixels < 0) {
        throw new IllegalArgumentException("widthPixels must not be negative");
      }
      if (heightPixels != null && heightPixels < 0) {
        throw new IllegalArgumentException("heightPixels must not be negative");
      }
      if (description != null && description.isBlank()) {
        throw new IllegalArgumentException("description must not be blank");
      }
    }
  }

  /** Factual chart report surfaced through the drawing-object inventory. */
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
  record Shape(
      String name,
      DrawingAnchorReport anchor,
      ExcelDrawingShapeKind kind,
      String presetGeometryToken,
      String text,
      int childCount)
      implements DrawingObjectReport {
    public Shape {
      validateCommon(name, anchor);
      Objects.requireNonNull(kind, "kind must not be null");
      if (presetGeometryToken != null && presetGeometryToken.isBlank()) {
        throw new IllegalArgumentException("presetGeometryToken must not be blank");
      }
      if (text != null && text.isBlank()) {
        throw new IllegalArgumentException("text must not be blank");
      }
      if (childCount < 0) {
        throw new IllegalArgumentException("childCount must not be negative");
      }
    }
  }

  /** Factual embedded-object report. */
  record EmbeddedObject(
      String name,
      DrawingAnchorReport anchor,
      ExcelEmbeddedObjectPackagingKind packagingKind,
      String label,
      String fileName,
      String command,
      String contentType,
      long byteSize,
      String sha256,
      ExcelPictureFormat previewFormat,
      Long previewByteSize,
      String previewSha256)
      implements DrawingObjectReport {
    public EmbeddedObject {
      validateCommon(name, anchor);
      Objects.requireNonNull(packagingKind, "packagingKind must not be null");
      contentType = requireNonBlank(contentType, "contentType");
      sha256 = requireNonBlank(sha256, "sha256");
      if (label != null && label.isBlank()) {
        throw new IllegalArgumentException("label must not be blank");
      }
      if (fileName != null && fileName.isBlank()) {
        throw new IllegalArgumentException("fileName must not be blank");
      }
      if (command != null && command.isBlank()) {
        throw new IllegalArgumentException("command must not be blank");
      }
      if (byteSize < 0L) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      if (previewFormat == null && previewByteSize != null) {
        throw new IllegalArgumentException("previewByteSize requires previewFormat");
      }
      if (previewSha256 != null && previewSha256.isBlank()) {
        throw new IllegalArgumentException("previewSha256 must not be blank");
      }
      if (previewFormat == null && previewSha256 != null) {
        throw new IllegalArgumentException("previewSha256 requires previewFormat");
      }
      if (previewByteSize != null && previewByteSize < 0L) {
        throw new IllegalArgumentException("previewByteSize must not be negative");
      }
    }
  }

  /** Factual signature-line report. */
  record SignatureLine(
      String name,
      DrawingAnchorReport anchor,
      String setupId,
      Boolean allowComments,
      String signingInstructions,
      String suggestedSigner,
      String suggestedSigner2,
      String suggestedSignerEmail,
      ExcelPictureFormat previewFormat,
      String previewContentType,
      Long previewByteSize,
      String previewSha256,
      Integer previewWidthPixels,
      Integer previewHeightPixels)
      implements DrawingObjectReport {
    public SignatureLine {
      validateCommon(name, anchor);
      if (setupId != null && setupId.isBlank()) {
        throw new IllegalArgumentException("setupId must not be blank");
      }
      if (signingInstructions != null && signingInstructions.isBlank()) {
        throw new IllegalArgumentException("signingInstructions must not be blank");
      }
      if (suggestedSigner != null && suggestedSigner.isBlank()) {
        throw new IllegalArgumentException("suggestedSigner must not be blank");
      }
      if (suggestedSigner2 != null && suggestedSigner2.isBlank()) {
        throw new IllegalArgumentException("suggestedSigner2 must not be blank");
      }
      if (suggestedSignerEmail != null && suggestedSignerEmail.isBlank()) {
        throw new IllegalArgumentException("suggestedSignerEmail must not be blank");
      }
      if (previewFormat == null && previewContentType != null) {
        throw new IllegalArgumentException("previewContentType requires previewFormat");
      }
      if (previewContentType != null && previewContentType.isBlank()) {
        throw new IllegalArgumentException("previewContentType must not be blank");
      }
      if (previewFormat == null && previewByteSize != null) {
        throw new IllegalArgumentException("previewByteSize requires previewFormat");
      }
      if (previewByteSize != null && previewByteSize < 0L) {
        throw new IllegalArgumentException("previewByteSize must not be negative");
      }
      if (previewFormat == null && previewSha256 != null) {
        throw new IllegalArgumentException("previewSha256 requires previewFormat");
      }
      if (previewSha256 != null && previewSha256.isBlank()) {
        throw new IllegalArgumentException("previewSha256 must not be blank");
      }
      if (previewWidthPixels != null && previewWidthPixels < 0) {
        throw new IllegalArgumentException("previewWidthPixels must not be negative");
      }
      if (previewHeightPixels != null && previewHeightPixels < 0) {
        throw new IllegalArgumentException("previewHeightPixels must not be negative");
      }
    }
  }

  private static void validateCommon(String name, DrawingAnchorReport anchor) {
    requireNonBlank(name, "name");
    Objects.requireNonNull(anchor, "anchor must not be null");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static List<String> copyNonBlankStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = new ArrayList<>(values.size());
    for (String value : values) {
      copy.add(requireNonBlank(value, fieldName + " value"));
    }
    return List.copyOf(copy);
  }
}
