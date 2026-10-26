package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;
import java.util.Optional;

/** Authoritative signature-line creation or replacement payload. */
@CatalogType(
    id = "SignatureLineSpec",
    summary =
        "Named signature-line authoring payload for SET_SIGNATURE_LINE. allowComments is explicit and plainSignature is optional, but caption or suggested signer metadata must still be present.",
    displayOrder = 38)
public record SignatureLineSpec(
    String name,
    DrawingAnchorSpec anchor,
    boolean allowComments,
    Optional<String> signingInstructions,
    Optional<String> suggestedSigner,
    Optional<String> suggestedSigner2,
    Optional<String> suggestedSignerEmail,
    Optional<String> caption,
    Optional<String> invalidStamp,
    Optional<PictureDataSpec> plainSignature) {

  public SignatureLineSpec {
    requireNonBlank(name, "name");
    anchor = requireTwoCellAnchor(anchor);
    signingInstructions = normalizeOptional(signingInstructions, "signingInstructions");
    suggestedSigner = normalizeOptional(suggestedSigner, "suggestedSigner");
    suggestedSigner2 = normalizeOptional(suggestedSigner2, "suggestedSigner2");
    suggestedSignerEmail = normalizeOptional(suggestedSignerEmail, "suggestedSignerEmail");
    caption = normalizeOptional(caption, "caption");
    invalidStamp = normalizeOptional(invalidStamp, "invalidStamp");
    Objects.requireNonNull(plainSignature, "plainSignature must not be null");
    if (caption.isPresent() && caption.orElseThrow().lines().count() > 3L) {
      throw new IllegalArgumentException("caption must contain at most three lines");
    }
    if (caption.isEmpty()
        && suggestedSigner.isEmpty()
        && suggestedSigner2.isEmpty()
        && suggestedSignerEmail.isEmpty()) {
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

  private static Optional<String> normalizeOptional(Optional<String> value, String fieldName) {
    Optional<String> normalized = Objects.requireNonNull(value, fieldName + " must not be null");
    if (normalized.isEmpty()) {
      return Optional.empty();
    }
    String presentValue = normalized.orElseThrow();
    if (presentValue.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return Optional.of(presentValue);
  }

  private static DrawingAnchorSpec requireTwoCellAnchor(DrawingAnchorSpec anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorSpec.TwoCell twoCell -> twoCell;
    };
  }
}
