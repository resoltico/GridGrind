package dev.erst.gridgrind.core.model.security;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.source.SecretReference;
import java.util.Objects;
import java.util.Optional;

/** Optional OOXML package-open settings for encrypted existing workbook sources. */
public record OoxmlOpenSecurityInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<SecretReference> passwordRef) {
  public OoxmlOpenSecurityInput {
    Objects.requireNonNull(passwordRef, "passwordRef must not be null");
  }
}
