package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;
import java.util.Optional;

/** Optional OOXML package-open settings for encrypted existing workbook sources. */
public record OoxmlOpenSecurityInput(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> password) {
  @JsonCreator
  static OoxmlOpenSecurityInput create(@JsonProperty("password") String password) {
    return new OoxmlOpenSecurityInput(Optional.ofNullable(password));
  }

  public OoxmlOpenSecurityInput {
    password = normalizeOptional(password, "password");
  }

  boolean isEmpty() {
    return password.isEmpty();
  }

  private static Optional<String> normalizeOptional(Optional<String> value, String fieldName) {
    Optional<String> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
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
