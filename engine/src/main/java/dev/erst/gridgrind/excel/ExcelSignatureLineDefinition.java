package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Authoritative signature-line creation or replacement payload. */
public record ExcelSignatureLineDefinition(
    String name,
    ExcelDrawingAnchor.TwoCell anchor,
    boolean allowComments,
    @Nullable String signingInstructions,
    @Nullable String suggestedSigner,
    @Nullable String suggestedSigner2,
    @Nullable String suggestedSignerEmail,
    @Nullable String caption,
    @Nullable String invalidStamp,
    Optional<ExcelPictureFormat> plainSignatureFormat,
    Optional<ExcelBinaryData> plainSignature) {
  public ExcelSignatureLineDefinition {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    Objects.requireNonNull(anchor, "anchor must not be null");
    Objects.requireNonNull(plainSignatureFormat, "plainSignatureFormat must not be null");
    Objects.requireNonNull(plainSignature, "plainSignature must not be null");
    signingInstructions =
        normalizeOptional(signingInstructions, "signingInstructions").orElse(null);
    suggestedSigner = normalizeOptional(suggestedSigner, "suggestedSigner").orElse(null);
    suggestedSigner2 = normalizeOptional(suggestedSigner2, "suggestedSigner2").orElse(null);
    suggestedSignerEmail =
        normalizeOptional(suggestedSignerEmail, "suggestedSignerEmail").orElse(null);
    caption = normalizeOptional(caption, "caption").orElse(null);
    invalidStamp = normalizeOptional(invalidStamp, "invalidStamp").orElse(null);
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
    if (plainSignature.isEmpty() && plainSignatureFormat.isPresent()) {
      throw new IllegalArgumentException("plainSignatureFormat requires plainSignature");
    }
    if (plainSignature.isPresent() && plainSignatureFormat.isEmpty()) {
      throw new IllegalArgumentException("plainSignature requires plainSignatureFormat");
    }
  }

  private static Optional<String> normalizeOptional(@Nullable String value, String fieldName) {
    if (value == null) {
      return Optional.empty();
    }
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return Optional.of(value);
  }
}
