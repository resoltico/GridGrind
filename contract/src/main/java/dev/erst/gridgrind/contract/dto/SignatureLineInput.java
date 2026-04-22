package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Authoritative signature-line creation or replacement payload. */
public record SignatureLineInput(
    String name,
    DrawingAnchorInput anchor,
    Boolean allowComments,
    String signingInstructions,
    String suggestedSigner,
    String suggestedSigner2,
    String suggestedSignerEmail,
    String caption,
    String invalidStamp,
    PictureDataInput plainSignature) {
  public SignatureLineInput {
    requireNonBlank(name, "name");
    anchor = requireTwoCellAnchor(anchor);
    allowComments = allowComments == null ? Boolean.TRUE : allowComments;
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
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
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

  private static DrawingAnchorInput requireTwoCellAnchor(DrawingAnchorInput anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell -> twoCell;
    };
  }
}
