package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import java.util.Objects;
import java.util.Optional;

/** Factual OOXML package-signature report for one signature part. */
public record OoxmlSignatureReport(
    String packagePartName,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<SignerIdentity> signer,
    ExcelOoxmlSignatureState state) {
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
    packagePartName = requireNonBlank(packagePartName, "packagePartName");
    signer = Objects.requireNonNullElseGet(signer, Optional::empty);
    Objects.requireNonNull(state, "state must not be null");
  }

  /** Signer identity material attached to one OOXML package signature. */
  public record SignerIdentity(String subject, String issuer, String serialNumberHex) {
    public SignerIdentity {
      subject = requireNonBlank(subject, "subject");
      issuer = requireNonBlank(issuer, "issuer");
      serialNumberHex = requireNonBlank(serialNumberHex, "serialNumberHex");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
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
