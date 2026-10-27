package dev.erst.gridgrind.core.model.security;

import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** OOXML package-signing settings applied during workbook persistence. */
public record ResolvedSignature(
    Path pkcs12Path,
    String keystorePassword,
    String keyPassword,
    Optional<String> alias,
    ExcelOoxmlSignatureDigestAlgorithm digestAlgorithm,
    Optional<String> description) {
  public ResolvedSignature {
    Objects.requireNonNull(pkcs12Path, "pkcs12Path must not be null");
    keystorePassword = TextValidation.requireNonBlank(keystorePassword, "keystorePassword");
    keyPassword =
        keyPassword == null
            ? keystorePassword
            : TextValidation.requireNonBlank(keyPassword, "keyPassword");
    alias = normalizeOptional(alias, "alias");
    digestAlgorithm =
        Objects.requireNonNullElse(digestAlgorithm, ExcelOoxmlSignatureDigestAlgorithm.SHA256);
    description = normalizeOptional(description, "description");
  }

  private static Optional<String> normalizeOptional(Optional<String> value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isEmpty()) {
      return Optional.empty();
    }
    if (value.orElseThrow().isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  @Override
  public String toString() {
    return "ResolvedSignature[pkcs12Path="
        + pkcs12Path
        + ", keystorePassword=redacted, keyPassword=redacted, alias="
        + alias
        + ", digestAlgorithm="
        + digestAlgorithm
        + ", description="
        + description
        + "]";
  }
}
