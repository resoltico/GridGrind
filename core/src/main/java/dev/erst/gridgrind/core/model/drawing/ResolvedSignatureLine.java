package dev.erst.gridgrind.core.model.drawing;

import static dev.erst.gridgrind.core.model.drawing.SignatureLineSpec.normalizeOptional;

import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Objects;
import java.util.Optional;

/** Signature-line payload with materialized preview bytes and explicit optional metadata. */
public record ResolvedSignatureLine(
    String name,
    DrawingAnchorSpec.TwoCell anchor,
    boolean allowComments,
    Optional<String> signingInstructions,
    Optional<String> suggestedSigner,
    Optional<String> suggestedSigner2,
    Optional<String> suggestedSignerEmail,
    Optional<String> caption,
    Optional<String> invalidStamp,
    Optional<ExcelPictureFormat> plainSignatureFormat,
    Optional<BinaryContent> plainSignature) {
  public ResolvedSignatureLine {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(plainSignatureFormat, "plainSignatureFormat must not be null");
    Objects.requireNonNull(plainSignature, "plainSignature must not be null");
    signingInstructions = normalizeOptional(signingInstructions, "signingInstructions");
    suggestedSigner = normalizeOptional(suggestedSigner, "suggestedSigner");
    suggestedSigner2 = normalizeOptional(suggestedSigner2, "suggestedSigner2");
    suggestedSignerEmail = normalizeOptional(suggestedSignerEmail, "suggestedSignerEmail");
    caption = normalizeOptional(caption, "caption");
    invalidStamp = normalizeOptional(invalidStamp, "invalidStamp");
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
    if (plainSignature.isEmpty() && plainSignatureFormat.isPresent()) {
      throw new IllegalArgumentException("plainSignatureFormat requires plainSignature");
    }
    if (plainSignature.isPresent() && plainSignatureFormat.isEmpty()) {
      throw new IllegalArgumentException("plainSignature requires plainSignatureFormat");
    }
  }
}
