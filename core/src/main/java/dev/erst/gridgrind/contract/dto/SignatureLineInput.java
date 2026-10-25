package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;
import java.util.Optional;

/** Authoritative signature-line creation or replacement payload. */
public record SignatureLineInput(
    String name,
    DrawingAnchorInput anchor,
    boolean allowComments,
    Optional<String> signingInstructions,
    Optional<String> suggestedSigner,
    Optional<String> suggestedSigner2,
    Optional<String> suggestedSignerEmail,
    Optional<String> caption,
    Optional<String> invalidStamp,
    Optional<PictureDataInput> plainSignature) {
  /** Creates a signature-line payload from the authored wire shape. */
  @JsonCreator
  @SuppressWarnings("PMD.ExcessiveParameterList")
  public SignatureLineInput(
      @JsonProperty("name") String name,
      @JsonProperty("anchor") DrawingAnchorInput anchor,
      @JsonProperty("allowComments") Boolean allowComments,
      @JsonProperty("signingInstructions") Optional<String> signingInstructions,
      @JsonProperty("suggestedSigner") Optional<String> suggestedSigner,
      @JsonProperty("suggestedSigner2") Optional<String> suggestedSigner2,
      @JsonProperty("suggestedSignerEmail") Optional<String> suggestedSignerEmail,
      @JsonProperty("caption") Optional<String> caption,
      @JsonProperty("invalidStamp") Optional<String> invalidStamp,
      @JsonProperty("plainSignature") Optional<PictureDataInput> plainSignature) {
    this(
        name,
        anchor,
        Objects.requireNonNull(allowComments, "allowComments must not be null").booleanValue(),
        Objects.requireNonNull(signingInstructions, "signingInstructions must not be null"),
        Objects.requireNonNull(suggestedSigner, "suggestedSigner must not be null"),
        Objects.requireNonNull(suggestedSigner2, "suggestedSigner2 must not be null"),
        Objects.requireNonNull(suggestedSignerEmail, "suggestedSignerEmail must not be null"),
        Objects.requireNonNull(caption, "caption must not be null"),
        Objects.requireNonNull(invalidStamp, "invalidStamp must not be null"),
        Objects.requireNonNull(plainSignature, "plainSignature must not be null"));
  }

  public SignatureLineInput {
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

  private static DrawingAnchorInput requireTwoCellAnchor(DrawingAnchorInput anchor) {
    Objects.requireNonNull(anchor, "anchor must not be null");
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell -> twoCell;
    };
  }
}
