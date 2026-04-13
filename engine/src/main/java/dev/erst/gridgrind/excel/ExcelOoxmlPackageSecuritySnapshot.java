package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Immutable factual snapshot of one workbook package's OOXML encryption and signature state. */
public record ExcelOoxmlPackageSecuritySnapshot(
    ExcelOoxmlEncryptionSnapshot encryption, List<ExcelOoxmlSignatureSnapshot> signatures) {
  public ExcelOoxmlPackageSecuritySnapshot {
    Objects.requireNonNull(encryption, "encryption must not be null");
    signatures = copyValues(signatures, "signatures");
  }

  /** Returns the canonical snapshot for a plain OOXML workbook with no package security. */
  public static ExcelOoxmlPackageSecuritySnapshot none() {
    return new ExcelOoxmlPackageSecuritySnapshot(ExcelOoxmlEncryptionSnapshot.none(), List.of());
  }

  /**
   * Returns whether the workbook package is encrypted or carries at least one package signature.
   */
  public boolean isSecure() {
    return encryption.encrypted() || !signatures.isEmpty();
  }

  ExcelOoxmlPackageSecuritySnapshot afterMutation() {
    if (signatures.isEmpty()) {
      return this;
    }
    return new ExcelOoxmlPackageSecuritySnapshot(
        encryption, signatures.stream().map(ExcelOoxmlSignatureSnapshot::afterMutation).toList());
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
