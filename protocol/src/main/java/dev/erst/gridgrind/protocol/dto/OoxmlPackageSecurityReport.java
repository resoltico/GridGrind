package dev.erst.gridgrind.protocol.dto;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Factual OOXML package-security report covering encryption and package signatures. */
public record OoxmlPackageSecurityReport(
    OoxmlEncryptionReport encryption, List<OoxmlSignatureReport> signatures) {
  public OoxmlPackageSecurityReport {
    Objects.requireNonNull(encryption, "encryption must not be null");
    signatures = copyValues(signatures, "signatures");
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
