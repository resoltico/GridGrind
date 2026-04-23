package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Objects;

/** Authoritative signature-line creation or replacement payload. */
public record ExcelSignatureLineDefinition(
    String name,
    ExcelDrawingAnchor.TwoCell anchor,
    boolean allowComments,
    String signingInstructions,
    String suggestedSigner,
    String suggestedSigner2,
    String suggestedSignerEmail,
    String caption,
    String invalidStamp,
    ExcelPictureFormat plainSignatureFormat,
    ExcelBinaryData plainSignature) {
  public ExcelSignatureLineDefinition {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(anchor, "anchor must not be null");
    signingInstructions = normalizeOptional(signingInstructions, "signingInstructions");
    suggestedSigner = normalizeOptional(suggestedSigner, "suggestedSigner");
    suggestedSigner2 = normalizeOptional(suggestedSigner2, "suggestedSigner2");
    suggestedSignerEmail = normalizeOptional(suggestedSignerEmail, "suggestedSignerEmail");
    caption = normalizeOptional(caption, "caption");
    invalidStamp = normalizeOptional(invalidStamp, "invalidStamp");
    if (caption != null && caption.lines().count() > 3L) {
      throw new IllegalArgumentException("caption must contain at most three lines");
    }
    if (caption == null
        && suggestedSigner == null
        && suggestedSigner2 == null
        && suggestedSignerEmail == null) {
      throw new IllegalArgumentException(
          "caption or at least one suggested signer field must be provided");
    }
    if (plainSignature == null && plainSignatureFormat != null) {
      throw new IllegalArgumentException("plainSignatureFormat requires plainSignature");
    }
    if (plainSignature != null && plainSignatureFormat == null) {
      throw new IllegalArgumentException("plainSignature requires plainSignatureFormat");
    }
  }

  private static String normalizeOptional(String value, String fieldName) {
    if (value == null) {
      return null;
    }
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
