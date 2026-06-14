package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Test fixture helpers for factual drawing-object reports. */
final class DrawingObjectReportTestSupport {
  private DrawingObjectReportTestSupport() {}

  static DrawingObjectReport.SignatureLine signatureLine(
      String name,
      DrawingAnchorReport anchor,
      Optional<DrawingObjectReport.SignatureSetup> setup,
      Optional<DrawingObjectReport.SignaturePreview> preview) {
    return new DrawingObjectReport.SignatureLine(
        name,
        anchor,
        Objects.requireNonNullElseGet(setup, Optional::empty),
        Objects.requireNonNullElseGet(preview, Optional::empty));
  }

  static Optional<DrawingObjectReport.SignatureSetup> signatureSetup(
      @Nullable String setupId,
      @Nullable Boolean allowComments,
      @Nullable String signingInstructions,
      @Nullable String suggestedSigner,
      @Nullable String suggestedSigner2,
      @Nullable String suggestedSignerEmail) {
    if (setupId == null
        && allowComments == null
        && signingInstructions == null
        && suggestedSigner == null
        && suggestedSigner2 == null
        && suggestedSignerEmail == null) {
      return Optional.empty();
    }
    return Optional.of(
        new DrawingObjectReport.SignatureSetup(
            Optional.ofNullable(setupId),
            Optional.ofNullable(allowComments),
            Optional.ofNullable(signingInstructions),
            Optional.ofNullable(suggestedSigner),
            Optional.ofNullable(suggestedSigner2),
            Optional.ofNullable(suggestedSignerEmail)));
  }

  static Optional<DrawingObjectReport.SignaturePreview> signaturePreview(
      @Nullable ExcelPictureFormat previewFormat,
      @Nullable String previewContentType,
      @Nullable Long previewByteSize,
      @Nullable String previewSha256,
      @Nullable Integer previewWidthPixels,
      @Nullable Integer previewHeightPixels) {
    if (previewFormat == null
        && previewContentType == null
        && previewByteSize == null
        && previewSha256 == null
        && previewWidthPixels == null
        && previewHeightPixels == null) {
      return Optional.empty();
    }
    ExcelPictureFormat format =
        requirePreviewFormat(
            previewFormat,
            previewContentType,
            previewByteSize,
            previewSha256,
            previewWidthPixels,
            previewHeightPixels);
    return Optional.of(
        new DrawingObjectReport.SignaturePreview(
            format,
            requirePreviewContentType(previewContentType),
            requirePreviewByteSize(previewByteSize),
            Optional.ofNullable(previewSha256),
            Optional.ofNullable(previewWidthPixels),
            Optional.ofNullable(previewHeightPixels)));
  }

  private static ExcelPictureFormat requirePreviewFormat(
      @Nullable ExcelPictureFormat previewFormat,
      @Nullable String previewContentType,
      @Nullable Long previewByteSize,
      @Nullable String previewSha256,
      @Nullable Integer previewWidthPixels,
      @Nullable Integer previewHeightPixels) {
    if (previewFormat != null) {
      return previewFormat;
    }
    if (previewContentType != null) {
      throw new IllegalArgumentException("previewContentType requires previewFormat");
    }
    if (previewByteSize != null) {
      throw new IllegalArgumentException("previewByteSize requires previewFormat");
    }
    if (previewSha256 != null) {
      throw new IllegalArgumentException("previewSha256 requires previewFormat");
    }
    if (previewWidthPixels != null) {
      throw new IllegalArgumentException("previewWidthPixels requires previewFormat");
    }
    if (previewHeightPixels != null) {
      throw new IllegalArgumentException("previewHeightPixels requires previewFormat");
    }
    throw new IllegalArgumentException("previewFormat is required when preview metadata exists");
  }

  private static String requirePreviewContentType(@Nullable String previewContentType) {
    if (previewContentType == null) {
      throw new IllegalArgumentException("previewContentType must not be null when preview exists");
    }
    return previewContentType;
  }

  private static long requirePreviewByteSize(@Nullable Long previewByteSize) {
    if (previewByteSize == null) {
      throw new IllegalArgumentException("previewByteSize must not be null when preview exists");
    }
    return previewByteSize;
  }
}
