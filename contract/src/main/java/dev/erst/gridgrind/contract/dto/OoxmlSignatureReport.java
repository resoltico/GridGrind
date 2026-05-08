package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import java.util.Objects;
import java.util.Optional;

/** Factual OOXML package-signature report for one signature part. */
public record OoxmlSignatureReport(
    String packagePartName,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> signerSubject,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> signerIssuer,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> serialNumberHex,
    ExcelOoxmlSignatureState state) {
  public OoxmlSignatureReport {
    packagePartName = requireNonBlank(packagePartName, "packagePartName");
    Objects.requireNonNull(signerSubject, "signerSubject must not be null");
    Objects.requireNonNull(signerIssuer, "signerIssuer must not be null");
    Objects.requireNonNull(serialNumberHex, "serialNumberHex must not be null");
    Objects.requireNonNull(state, "state must not be null");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
