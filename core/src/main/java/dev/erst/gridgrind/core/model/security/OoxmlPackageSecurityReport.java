package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Factual OOXML package-security report covering encryption and package signatures. */
@CatalogType(
    id = "OoxmlPackageSecurityReport",
    summary = "Factual OOXML package-security report covering encryption and package signatures.",
    displayOrder = 23)
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
