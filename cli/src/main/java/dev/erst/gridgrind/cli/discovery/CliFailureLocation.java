package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.json.PayloadException;
import java.util.Objects;
import java.util.Optional;

/** Compact JSON cursor metadata carried by CLI failure reports when one request location exists. */
public record CliFailureLocation(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> jsonPath,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> jsonLine,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> jsonColumn) {
  public CliFailureLocation {
    jsonPath = CliDiscoveryValidation.normalizeOptionalString(jsonPath, "jsonPath");
    jsonLine = normalizePositiveInteger(jsonLine, "jsonLine");
    jsonColumn = normalizePositiveInteger(jsonColumn, "jsonColumn");
    if (jsonLine.isPresent() != jsonColumn.isPresent()) {
      throw new IllegalArgumentException(
          "jsonLine and jsonColumn must either both be present or both be absent");
    }
  }

  /** Returns the empty location variant used when no JSON cursor is available. */
  public static CliFailureLocation unavailable() {
    return new CliFailureLocation(Optional.empty(), Optional.empty(), Optional.empty());
  }

  /** Extracts one normalized JSON cursor from a payload exception when present. */
  public static Optional<CliFailureLocation> from(Throwable exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    if (exception instanceof PayloadException payloadException) {
      return from(
          payloadException.jsonPath(), payloadException.jsonLine(), payloadException.jsonColumn());
    }
    return Optional.empty();
  }

  /** Builds one normalized location from explicit JSON path and cursor facts when available. */
  public static Optional<CliFailureLocation> from(
      Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
    CliFailureLocation location = new CliFailureLocation(jsonPath, jsonLine, jsonColumn);
    return location.isAvailable() ? Optional.of(location) : Optional.empty();
  }

  /** Returns whether this location carries at least one concrete request cursor fact. */
  @JsonIgnore
  public boolean isAvailable() {
    return jsonPath.isPresent() || jsonLine.isPresent();
  }

  private static Optional<Integer> normalizePositiveInteger(
      Optional<Integer> value, String fieldName) {
    Optional<Integer> normalized = Objects.requireNonNull(value, fieldName + " must not be null");
    if (normalized.isPresent() && normalized.orElseThrow() <= 0) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
    return normalized;
  }
}
