package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import java.util.Objects;

/** Factual OOXML package-signature report for one signature part. */
public record OoxmlSignatureReport(
    String packagePartName,
    String signerSubject,
    String signerIssuer,
    String serialNumberHex,
    ExcelOoxmlSignatureState state) {
  public OoxmlSignatureReport {
    packagePartName = requireNonBlank(packagePartName, "packagePartName");
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
