package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.Objects;
import java.util.Optional;

/** One stderr-only notice emitted when a requested response file falls back to standard output. */
public record CliTransportNotice(
    Reason reason,
    Destination wroteTo,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> responsePath) {
  public CliTransportNotice {
    Objects.requireNonNull(reason, "reason must not be null");
    Objects.requireNonNull(wroteTo, "wroteTo must not be null");
    responsePath = Objects.requireNonNullElseGet(responsePath, Optional::empty);
    responsePath =
        responsePath.map(path -> CliDiscoveryValidation.requireNonBlank(path, "responsePath"));
  }

  /** Creates the fallback notice for a response file that could not be written. */
  public static CliTransportNotice stdoutFallback(String responsePath) {
    return new CliTransportNotice(
        Reason.RESPONSE_WRITE_FAILED, Destination.STDOUT, Optional.of(responsePath));
  }

  /** Creates the no-execution notice for a response destination that could not be reserved. */
  public static CliTransportNotice reservationFailure(Reason reason, String responsePath) {
    return new CliTransportNotice(reason, Destination.NOT_DELIVERED, Optional.of(responsePath));
  }

  /** Stable reason for response-file transport behavior. */
  public enum Reason {
    RESPONSE_PATH_EXISTS,
    RESPONSE_PATH_DIRECTORY,
    RESPONSE_PATH_UNWRITABLE,
    RESPONSE_WRITE_FAILED
  }

  /** Destination used for the primary payload after transport routing. */
  public enum Destination {
    STDOUT,
    NOT_DELIVERED
  }
}
