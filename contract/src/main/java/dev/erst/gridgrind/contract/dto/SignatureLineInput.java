package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import java.util.Objects;
import java.util.Optional;

/** Authoritative signature-line creation or replacement payload. */
public record SignatureLineInput(
    String name,
    DrawingAnchorInput anchor,
    Boolean allowComments,
    Optional<String> signingInstructions,
    Optional<String> suggestedSigner,
    Optional<String> suggestedSigner2,
    Optional<String> suggestedSignerEmail,
    Optional<String> caption,
    Optional<String> invalidStamp,
    Optional<PictureDataInput> plainSignature) {
  @JsonCreator(mode = JsonCreator.Mode.DELEGATING)
  static SignatureLineInput create(SignatureLineJson json) {
    return new SignatureLineInput(
        json.name(),
        json.anchor(),
        json.allowComments() == null ? Boolean.TRUE : json.allowComments(),
        json.signingInstructions(),
        json.suggestedSigner(),
        json.suggestedSigner2(),
        json.suggestedSignerEmail(),
        json.caption(),
        json.invalidStamp(),
        json.plainSignature());
  }

  public SignatureLineInput {
    requireNonBlank(name, "name");
    anchor = requireTwoCellAnchor(anchor);
    Objects.requireNonNull(allowComments, "allowComments must not be null");
    signingInstructions = normalizeOptional(signingInstructions, "signingInstructions");
    suggestedSigner = normalizeOptional(suggestedSigner, "suggestedSigner");
    suggestedSigner2 = normalizeOptional(suggestedSigner2, "suggestedSigner2");
    suggestedSignerEmail = normalizeOptional(suggestedSignerEmail, "suggestedSignerEmail");
    caption = normalizeOptional(caption, "caption");
    invalidStamp = normalizeOptional(invalidStamp, "invalidStamp");
    plainSignature = Objects.requireNonNullElseGet(plainSignature, Optional::empty);
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
    Optional<String> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    if (normalized.isEmpty()) {
      return Optional.empty();
    }
    String presentValue = normalized.orElseThrow();
    if (presentValue.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return Optional.of(presentValue);
  }

  private static DrawingAnchorInput requireTwoCellAnchor(DrawingAnchorInput anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell -> twoCell;
    };
  }

  private record SignatureLineJson(
      String name,
      DrawingAnchorInput anchor,
      Boolean allowComments,
      Optional<String> signingInstructions,
      Optional<String> suggestedSigner,
      Optional<String> suggestedSigner2,
      Optional<String> suggestedSignerEmail,
      Optional<String> caption,
      Optional<String> invalidStamp,
      Optional<PictureDataInput> plainSignature) {}
}
