package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import java.util.Objects;
import java.util.Optional;

/** OOXML package-signing settings applied during workbook persistence. */
@CatalogType(
    id = "OoxmlSignatureSpec",
    summary =
        "OOXML package-signing settings for workbook persistence. keyPassword defaults to keystorePassword and digestAlgorithm defaults to SHA256 when omitted. alias may be omitted only when the keystore contains exactly one signable private-key entry.",
    noteRefs = {"requestOwnedPathRule"},
    displayOrder = 22)
public record OoxmlSignatureSpec(
    String pkcs12Path,
    SecretReference keystorePasswordRef,
    Optional<SecretReference> keyPasswordRef,
    Optional<String> alias,
    @ProtocolField(optional = true) @WireDefault("SHA256")
        ExcelOoxmlSignatureDigestAlgorithm digestAlgorithm,
    Optional<String> description) {

  public OoxmlSignatureSpec {
    pkcs12Path = normalizeRequired(pkcs12Path, "pkcs12Path");
    Objects.requireNonNull(keystorePasswordRef, "keystorePasswordRef must not be null");
    Objects.requireNonNull(keyPasswordRef, "keyPasswordRef must not be null");
    alias = normalizeOptional(alias, "alias");
    Objects.requireNonNull(digestAlgorithm, "digestAlgorithm must not be null");
    description = normalizeOptional(description, "description");
  }

  /** Creates one signature payload whose key password matches the keystore password. */
  public static OoxmlSignatureSpec sameKeyPassword(
      String pkcs12Path,
      SecretReference keystorePasswordRef,
      Optional<String> alias,
      ExcelOoxmlSignatureDigestAlgorithm digestAlgorithm,
      Optional<String> description) {
    return new OoxmlSignatureSpec(
        pkcs12Path, keystorePasswordRef, Optional.empty(), alias, digestAlgorithm, description);
  }

  private static String normalizeRequired(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static Optional<String> normalizeOptional(Optional<String> value, String fieldName) {
    Optional<String> normalized = Objects.requireNonNull(value, fieldName + " must not be null");
    if (normalized.isEmpty()) {
      return Optional.empty();
    }
    String presentValue = normalized.orElseThrow();
    if (presentValue.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return Optional.of(presentValue);
  }
}
