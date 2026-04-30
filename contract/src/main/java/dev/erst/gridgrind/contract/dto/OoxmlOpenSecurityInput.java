package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** Optional OOXML package-open settings for encrypted existing workbook sources. */
public record OoxmlOpenSecurityInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> password) {
  public OoxmlOpenSecurityInput {
    password = normalizeOptional(password, "password");
  }

  boolean isEmpty() {
    return password.isEmpty();
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
