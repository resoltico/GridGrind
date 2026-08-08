package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** One stderr-only notice emitted when a requested response file falls back to standard output. */
public record CliTransportNotice(
    Destination wroteTo,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> responsePath) {
  public CliTransportNotice {
    Objects.requireNonNull(wroteTo, "wroteTo must not be null");
    responsePath = Objects.requireNonNullElseGet(responsePath, Optional::empty);
    responsePath =
        responsePath.map(path -> CliDiscoveryValidation.requireNonBlank(path, "responsePath"));
  }

  /** Creates the fallback notice for a response file that could not be written. */
  public static CliTransportNotice stdoutFallback(String responsePath) {
    return new CliTransportNotice(Destination.STDOUT, Optional.of(responsePath));
  }

  /** Destination used for the primary payload after transport routing. */
  public enum Destination {
    STDOUT
  }
}
