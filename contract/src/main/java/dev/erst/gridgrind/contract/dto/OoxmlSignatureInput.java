package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import java.util.Objects;
import java.util.Optional;

/** OOXML package-signing settings applied during workbook persistence. */
public record OoxmlSignatureInput(
    String pkcs12Path,
    String keystorePassword,
    String keyPassword,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> alias,
    ExcelOoxmlSignatureDigestAlgorithm digestAlgorithm,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> description) {
  @JsonCreator
  static OoxmlSignatureInput create(
      @JsonProperty("pkcs12Path") String pkcs12Path,
      @JsonProperty("keystorePassword") String keystorePassword,
      @JsonProperty("keyPassword") String keyPassword,
      @JsonProperty("alias") Optional<String> alias,
      @JsonProperty("digestAlgorithm") ExcelOoxmlSignatureDigestAlgorithm digestAlgorithm,
      @JsonProperty("description") Optional<String> description) {
    return new OoxmlSignatureInput(
        pkcs12Path, keystorePassword, keyPassword, alias, digestAlgorithm, description);
  }

  public OoxmlSignatureInput {
    pkcs12Path = normalizeRequired(pkcs12Path, "pkcs12Path");
    keystorePassword = normalizeRequired(keystorePassword, "keystorePassword");
    keyPassword = normalizeRequired(keyPassword, "keyPassword");
    alias = normalizeOptional(alias, "alias");
    Objects.requireNonNull(digestAlgorithm, "digestAlgorithm must not be null");
    description = normalizeOptional(description, "description");
  }

  /** Creates one signature payload whose key password matches the keystore password. */
  public static OoxmlSignatureInput sameKeyPassword(
      String pkcs12Path,
      String keystorePassword,
      Optional<String> alias,
      ExcelOoxmlSignatureDigestAlgorithm digestAlgorithm,
      Optional<String> description) {
    return new OoxmlSignatureInput(
        pkcs12Path, keystorePassword, keystorePassword, alias, digestAlgorithm, description);
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
