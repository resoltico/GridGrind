package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import java.util.Objects;

/** Immutable factual snapshot of one OOXML package-signature part. */
public record ExcelOoxmlSignatureSnapshot(
    String packagePartName,
    String signerSubject,
    String signerIssuer,
    String serialNumberHex,
    ExcelOoxmlSignatureState state) {
  public ExcelOoxmlSignatureSnapshot {
    packagePartName = requireNonBlank(packagePartName, "packagePartName");
    Objects.requireNonNull(state, "state must not be null");
  }

  ExcelOoxmlSignatureSnapshot afterMutation() {
    return new ExcelOoxmlSignatureSnapshot(
        packagePartName, signerSubject, signerIssuer, serialNumberHex, state.afterMutation());
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
