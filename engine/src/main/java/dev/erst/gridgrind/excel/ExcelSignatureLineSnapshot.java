package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Objects;

/** Immutable factual signature-line snapshot returned by workbook reads. */
public record ExcelSignatureLineSnapshot(
    String name,
    ExcelDrawingAnchor.TwoCell anchor,
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
    Integer previewHeightPixels) {
  public ExcelSignatureLineSnapshot {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(anchor, "anchor must not be null");
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
