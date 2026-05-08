package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import java.util.Objects;
import java.util.Optional;

/** Immutable factual snapshot of one OOXML package-signature part. */
public record ExcelOoxmlSignatureSnapshot(
    String packagePartName,
    Optional<String> signerSubject,
    Optional<String> signerIssuer,
    Optional<String> serialNumberHex,
    ExcelOoxmlSignatureState state) {
  public ExcelOoxmlSignatureSnapshot {
    packagePartName = requireNonBlank(packagePartName, "packagePartName");
    Objects.requireNonNull(signerSubject, "signerSubject must not be null");
    Objects.requireNonNull(signerIssuer, "signerIssuer must not be null");
    Objects.requireNonNull(serialNumberHex, "serialNumberHex must not be null");
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
