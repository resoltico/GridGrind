package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.Optional;

/** Factual OOXML package-signature report for one signature part. */
@CatalogType(
    id = "OoxmlSignatureReport",
    summary =
        "Factual OOXML package-signature report for one signature part. state reflects the currently loaded workbook package, including INVALIDATED_BY_MUTATION for source signatures after in-memory edits.",
    displayOrder = 24)
public record OoxmlSignatureReport(
    String packagePartName, Optional<SignerIdentity> signer, ExcelOoxmlSignatureState state) {
  /** Creates a signature report from flat signer facts while preserving grouped signer identity. */
  public OoxmlSignatureReport(
      String packagePartName,
      Optional<String> signerSubject,
      Optional<String> signerIssuer,
      Optional<String> serialNumberHex,
      ExcelOoxmlSignatureState state) {
    this(packagePartName, collapseSigner(signerSubject, signerIssuer, serialNumberHex), state);
  }

  public OoxmlSignatureReport {
    packagePartName = TextValidation.requireNonBlank(packagePartName, "packagePartName");
    signer = Objects.requireNonNullElseGet(signer, Optional::empty);
    Objects.requireNonNull(state, "state must not be null");
  }

  /** Signer identity material attached to one OOXML package signature. */
  @CatalogType(
      id = "OoxmlSignatureSignerIdentity",
      summary = "Signer identity material attached to one OOXML package signature report.",
      displayOrder = 25)
  public record SignerIdentity(String subject, String issuer, String serialNumberHex) {
    public SignerIdentity {
      subject = TextValidation.requireNonBlank(subject, "subject");
      issuer = TextValidation.requireNonBlank(issuer, "issuer");
      serialNumberHex = TextValidation.requireNonBlank(serialNumberHex, "serialNumberHex");
    }
  }

  private static Optional<SignerIdentity> collapseSigner(
      Optional<String> signerSubject,
      Optional<String> signerIssuer,
      Optional<String> serialNumberHex) {
    Optional<String> subject = Objects.requireNonNullElseGet(signerSubject, Optional::empty);
    Optional<String> issuer = Objects.requireNonNullElseGet(signerIssuer, Optional::empty);
    Optional<String> serial = Objects.requireNonNullElseGet(serialNumberHex, Optional::empty);
    int presentCount =
        (subject.isPresent() ? 1 : 0) + (issuer.isPresent() ? 1 : 0) + (serial.isPresent() ? 1 : 0);
    if (presentCount == 0) {
      return Optional.empty();
    }
    if (presentCount != 3) {
      throw new IllegalArgumentException(
          "signer identity must be either wholly absent or wholly present");
    }
    return Optional.of(
        new SignerIdentity(subject.orElseThrow(), issuer.orElseThrow(), serial.orElseThrow()));
  }
}
