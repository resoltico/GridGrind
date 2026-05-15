package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.json.PayloadException;
import java.util.Objects;
import java.util.Optional;

/** Compact JSON cursor metadata carried by CLI failure reports when one request location exists. */
public record CliFailureLocation(
    Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
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
  public static CliFailureLocation from(Throwable exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    if (exception instanceof PayloadException payloadException) {
      return new CliFailureLocation(
          payloadException.jsonPath(), payloadException.jsonLine(), payloadException.jsonColumn());
    }
    return unavailable();
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
