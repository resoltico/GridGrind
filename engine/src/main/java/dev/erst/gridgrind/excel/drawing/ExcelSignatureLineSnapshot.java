package dev.erst.gridgrind.excel.drawing;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Objects;
import java.util.Optional;

/** Immutable factual signature-line snapshot returned by workbook reads. */
public record ExcelSignatureLineSnapshot(
    String name,
    ExcelDrawingAnchor.TwoCell anchor,
    Optional<Setup> setup,
    Optional<Preview> preview) {
  public ExcelSignatureLineSnapshot {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(anchor, "anchor must not be null");
    setup = normalizeOptional(setup, "setup");
    preview = normalizeOptional(preview, "preview");
  }

  /** Optional signature-line setup metadata surfaced when Excel authored it. */
  public record Setup(
      Optional<String> setupId,
      Optional<Boolean> allowComments,
      Optional<String> signingInstructions,
      Optional<String> suggestedSigner,
      Optional<String> suggestedSigner2,
      Optional<String> suggestedSignerEmail) {
    public Setup {
      setupId = normalizeOptionalText(setupId, "setupId");
      allowComments = normalizeOptional(allowComments, "allowComments");
      signingInstructions = normalizeOptionalText(signingInstructions, "signingInstructions");
      suggestedSigner = normalizeOptionalText(suggestedSigner, "suggestedSigner");
      suggestedSigner2 = normalizeOptionalText(suggestedSigner2, "suggestedSigner2");
      suggestedSignerEmail = normalizeOptionalText(suggestedSignerEmail, "suggestedSignerEmail");
    }
  }

  /** Optional signature-line preview image metadata. */
  public record Preview(
      ExcelPictureFormat format,
      String contentType,
      long byteSize,
      Optional<String> sha256,
      Optional<Integer> widthPixels,
      Optional<Integer> heightPixels) {
    public Preview {
      Objects.requireNonNull(format, "format must not be null");
      contentType = requireNonBlank(contentType, "contentType");
      if (byteSize < 0L) {
        throw new IllegalArgumentException("byteSize must not be negative");
      }
      sha256 = normalizeOptionalText(sha256, "previewSha256");
      widthPixels = normalizeOptional(widthPixels, "widthPixels");
      heightPixels = normalizeOptional(heightPixels, "heightPixels");
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

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static Optional<String> normalizeOptionalText(Optional<String> value, String fieldName) {
    Optional<String> normalized = normalizeOptional(value, fieldName);
    return normalized.map(text -> requireNonBlank(text, fieldName));
  }

  private static <T> Optional<T> normalizeOptional(Optional<T> value, String fieldName) {
    Optional<T> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    normalized.ifPresent(
        entry -> Objects.requireNonNull(entry, fieldName + " must not contain null"));
    return normalized;
  }
}
