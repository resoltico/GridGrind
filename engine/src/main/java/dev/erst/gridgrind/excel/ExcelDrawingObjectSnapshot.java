package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.List;
import java.util.Objects;

/** Immutable factual drawing-object snapshot returned by workbook reads. */
public sealed interface ExcelDrawingObjectSnapshot
    permits ExcelDrawingObjectSnapshot.Picture,
        ExcelDrawingObjectSnapshot.Chart,
        ExcelDrawingObjectSnapshot.Shape,
        ExcelDrawingObjectSnapshot.EmbeddedObject,
        ExcelDrawingObjectSnapshot.SignatureLine {

  /** Sheet-local drawing object name. */
  String name();

  /** Stored anchor backing the drawing object. */
  ExcelDrawingAnchor anchor();

  /** Factual picture snapshot. */
  record Picture(
      String name,
      ExcelDrawingAnchor anchor,
      ExcelPictureFormat format,
      String contentType,
      long byteSize,
      String sha256,
      Integer widthPixels,
      Integer heightPixels,
      String description)
      implements ExcelDrawingObjectSnapshot {
    public Picture {
      validateCommon(name, anchor);
      Objects.requireNonNull(format, "format must not be null");
      contentType = requireNonBlank(contentType, "contentType");
      if (byteSize < 0L) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      sha256 = requireNonBlank(sha256, "sha256");
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

  /** Factual chart snapshot surfaced through the drawing-object inventory. */
  record Chart(
      String name,
      ExcelDrawingAnchor anchor,
      boolean supported,
      List<String> plotTypeTokens,
      String title)
      implements ExcelDrawingObjectSnapshot {
    public Chart {
      validateCommon(name, anchor);
      plotTypeTokens =
          List.copyOf(Objects.requireNonNull(plotTypeTokens, "plotTypeTokens must not be null"));
      for (String plotTypeToken : plotTypeTokens) {
        requireNonBlank(plotTypeToken, "plotTypeTokens value");
      }
      Objects.requireNonNull(title, "title must not be null");
    }
  }

  /** Factual non-binary drawing shape snapshot. */
  record Shape(
      String name,
      ExcelDrawingAnchor anchor,
      ExcelDrawingShapeKind kind,
      String presetGeometryToken,
      String text,
      int childCount)
      implements ExcelDrawingObjectSnapshot {
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

  /** Factual embedded-object snapshot. */
  record EmbeddedObject(
      String name,
      ExcelDrawingAnchor anchor,
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
      implements ExcelDrawingObjectSnapshot {
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
      if (previewByteSize != null && previewByteSize < 0L) {
        throw new IllegalArgumentException("previewByteSize must not be negative");
      }
      if (previewFormat == null && previewSha256 != null) {
        throw new IllegalArgumentException("previewSha256 requires previewFormat");
      }
    }
  }

  /** Factual signature-line snapshot. */
  record SignatureLine(
      String name,
      ExcelDrawingAnchor anchor,
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
      implements ExcelDrawingObjectSnapshot {
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

  private static void validateCommon(String name, ExcelDrawingAnchor anchor) {
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
}
