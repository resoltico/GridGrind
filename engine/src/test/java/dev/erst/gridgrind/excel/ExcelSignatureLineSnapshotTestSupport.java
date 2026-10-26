package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Shared test-only factories for signature-line snapshot fixtures. */
final class ExcelSignatureLineSnapshotTestSupport {
  private ExcelSignatureLineSnapshotTestSupport() {}

  @SuppressWarnings("PMD.ExcessiveParameterList")
  static DrawingObjectReport.SignatureLine signatureLineSnapshot(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
      @Nullable String setupId,
      @Nullable Boolean allowComments,
      @Nullable String signingInstructions,
      @Nullable String suggestedSigner,
      @Nullable String suggestedSigner2,
      @Nullable String suggestedSignerEmail,
      @Nullable ExcelPictureFormat previewFormat,
      @Nullable String previewContentType,
      @Nullable Long previewByteSize,
      @Nullable String previewSha256,
      @Nullable Integer previewWidthPixels,
      @Nullable Integer previewHeightPixels) {
    return new DrawingObjectReport.SignatureLine(
        name,
        DrawingGeometryFixtures.report(anchor),
        signatureSetup(
            setupId,
            allowComments,
            signingInstructions,
            suggestedSigner,
            suggestedSigner2,
            suggestedSignerEmail),
        signaturePreview(
            previewFormat,
            previewContentType,
            previewByteSize,
            previewSha256,
            previewWidthPixels,
            previewHeightPixels));
  }

  @SuppressWarnings("PMD.ExcessiveParameterList")
  static DrawingObjectReport.SignatureLine drawingSignatureLine(
      String name,
      DrawingAnchorSpec.TwoCell anchor,
      @Nullable String setupId,
      @Nullable Boolean allowComments,
      @Nullable String signingInstructions,
      @Nullable String suggestedSigner,
      @Nullable String suggestedSigner2,
      @Nullable String suggestedSignerEmail,
      @Nullable ExcelPictureFormat previewFormat,
      @Nullable String previewContentType,
      @Nullable Long previewByteSize,
      @Nullable String previewSha256,
      @Nullable Integer previewWidthPixels,
      @Nullable Integer previewHeightPixels) {
    return new DrawingObjectReport.SignatureLine(
        name,
        DrawingGeometryFixtures.report(anchor),
        signatureSetup(
            setupId,
            allowComments,
            signingInstructions,
            suggestedSigner,
            suggestedSigner2,
            suggestedSignerEmail),
        signaturePreview(
            previewFormat,
            previewContentType,
            previewByteSize,
            previewSha256,
            previewWidthPixels,
            previewHeightPixels));
  }

  private static Optional<DrawingObjectReport.SignatureSetup> signatureSetup(
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

  private static Optional<DrawingObjectReport.SignaturePreview> signaturePreview(
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
    if (previewFormat == null) {
      throw new IllegalArgumentException("previewFormat is required when preview metadata exists");
    }
    if (previewContentType == null) {
      throw new IllegalArgumentException("previewContentType must not be null when preview exists");
    }
    if (previewByteSize == null) {
      throw new IllegalArgumentException("previewByteSize must not be null when preview exists");
    }
    return Optional.of(
        new DrawingObjectReport.SignaturePreview(
            previewFormat,
            previewContentType,
            previewByteSize,
            Optional.ofNullable(previewSha256),
            Optional.ofNullable(previewWidthPixels),
            Optional.ofNullable(previewHeightPixels)));
  }
}
