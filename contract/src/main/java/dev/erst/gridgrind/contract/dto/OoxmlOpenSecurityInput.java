package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** Optional OOXML package-open settings for encrypted existing workbook sources. */
public record OoxmlOpenSecurityInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<SecretReference> passwordRef) {
  public OoxmlOpenSecurityInput {
    Objects.requireNonNull(passwordRef, "passwordRef must not be null");
  }

  boolean isEmpty() {
    return passwordRef.isEmpty();
  }
}
